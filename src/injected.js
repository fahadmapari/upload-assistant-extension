// Injected into the admin page via chrome.scripting.executeScript({ files: [...] })
// Tour data is pre-loaded into window.__tourExtData by popup.js before this runs.
(async function () {
  const tour = window.__tourExtData;
  const filled = [];
  const failed = [];
  const errors = {};

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  function setNativeInput(el, value) {
    if (!el) return false;
    try {
      const proto =
        el.tagName === "TEXTAREA"
          ? window.HTMLTextAreaElement.prototype
          : window.HTMLInputElement.prototype;
      const setter = Object.getOwnPropertyDescriptor(proto, "value").set;
      setter.call(el, value);
      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
      el.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
      return true;
    } catch {
      return false;
    }
  }

  async function ensurePanelOpen(el) {
    const panel = el.closest("mat-expansion-panel");
    if (!panel) return;
    const header = panel.querySelector("mat-expansion-panel-header");
    if (!header) return;
    if (header.getAttribute("aria-expanded") !== "true") {
      header.click();
      await sleep(600);
    }
  }

  async function fillNgbDatepicker(controlName, dateStr) {
    if (!dateStr) return;
    const [mm, dd, yyyy] = dateStr.split("/").map(Number);
    const input = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!input) { failed.push(controlName); return; }
    try {
      await ensurePanelOpen(input);
      input.click();
      await sleep(500);

      // Datepicker uses container="body" so it appends to <body>
      const picker = document.querySelector("ngb-datepicker.ngb-dp-body, ngb-datepicker.show");
      if (!picker) {
        errors[controlName] = "Datepicker popup did not open";
        failed.push(controlName);
        return;
      }

      const yearSel = picker.querySelector("select[aria-label='Select year']");
      const monthSel = picker.querySelector("select[aria-label='Select month']");

      // Set year first (affects which months are selectable)
      if (yearSel && yearSel.value !== String(yyyy)) {
        yearSel.value = String(yyyy);
        yearSel.dispatchEvent(new Event("change", { bubbles: true }));
        await sleep(300);
      }
      // Then set month
      if (monthSel && monthSel.value !== String(mm)) {
        monthSel.value = String(mm);
        monthSel.dispatchEvent(new Event("change", { bubbles: true }));
        await sleep(300);
      }

      // Click the matching day — skip "outside" (adjacent-month) cells
      const dayEls = picker.querySelectorAll(".ngb-dp-day:not(.disabled)");
      let clicked = false;
      for (const dayEl of dayEls) {
        const inner = dayEl.querySelector("[ngbdatepickerdayview]");
        if (!inner || inner.classList.contains("outside")) continue;
        if (inner.textContent.trim() === String(dd)) {
          dayEl.click();
          clicked = true;
          break;
        }
      }

      if (clicked) {
        filled.push(controlName);
      } else {
        errors[controlName] = `Day ${dd} not found in picker (month ${mm}/${yyyy})`;
        document.body.click();
        failed.push(controlName);
      }
      await sleep(200);
    } catch (e) {
      errors[controlName] = e.message;
      failed.push(controlName);
    }
  }

  async function fillByControl(controlName, value, reportKey) {
    const key = reportKey || controlName;
    if (!value && value !== 0) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el || el.disabled || el.getAttribute("readonly") === "readonly") return;
    await ensurePanelOpen(el);
    const tag = el.tagName.toLowerCase();
    if (tag === "input" || tag === "textarea") {
      if (setNativeInput(el, value)) filled.push(key);
      else failed.push(key);
    }
  }

  // Retrieve the ng-select Angular component instance for a specific host element.
  // el.__ngContext__ is the PARENT form's LView (shared by all ng-selects), so we must
  // match the found component back to `el` via its host-element getter/property.
  function getNgSelectComp(el) {
    // 1. Try Angular's official debug API (works in dev builds and Ivy production)
    try {
      if (typeof ng !== "undefined" && typeof ng.getComponent === "function") {
        const c = ng.getComponent(el);
        if (c && typeof c.open === "function" && c.itemsList) return c;
      }
    } catch (_) {}

    // 2. Search the parent LView for the component whose host element is `el`
    const lView = el.__ngContext__;
    if (!lView || !Array.isArray(lView)) return null;
    for (const item of lView) {
      if (
        item &&
        typeof item.open === "function" &&
        item.itemsList &&
        typeof item.select === "function"
      ) {
        // Verify this instance belongs to our element, not another ng-select
        const hostEl =
          item.element ||                        // ng-select public .element getter
          item._elementRef?.nativeElement ||     // common Angular DI pattern
          item.elementRef?.nativeElement;        // alternative accessor
        if (hostEl === el) return item;
      }
    }
    return null;
  }

  async function fillNgSelect(controlName, value, reportKey) {
    const key = reportKey || controlName;
    if (!value) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el) { failed.push(key); return; }
    try {
      await ensurePanelOpen(el);
      const comp = getNgSelectComp(el);
      if (comp) {
        // Primary path: use the ng-select component API directly
        comp.open();
        await sleep(500);
        const items = comp.itemsList.items || [];
        const match =
          items.find((i) => i.label === value) ||
          items.find((i) => i.label && i.label.toLowerCase().includes(value.toLowerCase()));
        if (match) {
          comp.select(match);
          comp.close();
          filled.push(key);
        } else {
          const available = items.map((i) => i.label).join(", ") || "(no items loaded)";
          errors[key] = `No match for "${value}". Available: ${available.slice(0, 120)}`;
          comp.close();
          failed.push(key);
        }
      } else {
        // Fallback: DOM click approach (for non-Ivy or unrecognised components)
        errors[key] = "Angular context not found — used DOM fallback";
        const trigger = el.querySelector(".ng-select-container") || el;
        trigger.click();
        await sleep(300);
        const searchInput = el.querySelector(".ng-input input");
        if (searchInput && !searchInput.readOnly && !searchInput.disabled) {
          setNativeInput(searchInput, value);
          await sleep(400);
        }
        let options = [];
        for (let i = 0; i < 10; i++) {
          options = [...document.querySelectorAll(".ng-option:not(.ng-option-disabled)")];
          if (options.length > 0) break;
          await sleep(200);
        }
        const match =
          options.find((o) => o.textContent.trim() === value) ||
          options.find((o) => o.textContent.trim().toLowerCase().includes(value.toLowerCase()));
        if (match) { delete errors[key]; match.click(); filled.push(key); }
        else {
          errors[key] = `DOM fallback: no option matching "${value}"`;
          document.body.click(); failed.push(key);
        }
      }
      await sleep(200);
    } catch (e) {
      errors[key] = e.message;
      console.error("[TourExt] fillNgSelect failed for", controlName, ":", e.message);
      failed.push(key);
    }
  }

  async function fillNgSelectMultiple(controlName, values, reportKey) {
    const key = reportKey || controlName;
    if (!values || !values.length) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el) { failed.push(key); return; }
    try {
      await ensurePanelOpen(el);
      const comp = getNgSelectComp(el);
      if (comp) {
        comp.open();
        await sleep(500);
        const items = comp.itemsList.items || [];
        for (const value of values) {
          const match =
            items.find((i) => i.label === value) ||
            items.find((i) => i.label && i.label.toLowerCase().includes(value.toLowerCase()));
          if (match && !match.selected) {
            comp.select(match);
            await sleep(50);
          }
        }
        comp.close();
        filled.push(key);
      } else {
        // Fallback: DOM click approach
        for (const value of values) {
          el.click();
          await sleep(200);
          const searchInput = el.querySelector(".ng-input input");
          if (searchInput && !searchInput.readOnly && !searchInput.disabled) {
            setNativeInput(searchInput, value);
            await sleep(300);
          }
          const options = [...document.querySelectorAll(".ng-option:not(.ng-option-disabled)")];
          const match =
            options.find((o) => o.textContent.trim() === value) ||
            options.find((o) => o.textContent.trim().toLowerCase().includes(value.toLowerCase()));
          if (match) { match.click(); await sleep(100); }
          else { document.body.click(); await sleep(100); }
        }
        filled.push(key);
      }
    } catch (e) {
      errors[key] = e.message;
      console.error("[TourExt] fillNgSelectMultiple failed for", controlName, ":", e.message);
      failed.push(key);
    }
  }

  function fillQuill(value) {
    if (!value) return;
    const editor = document.querySelector("quill-editor .ql-editor");
    if (!editor) {
      failed.push("description");
      return;
    }
    try {
      editor.focus();
      editor.innerHTML = `<p>${value.replace(/\n{2,}/g, "\n").replace(/\n/g, "</p><p>")}</p>`;
      editor.dispatchEvent(new Event("input", { bubbles: true }));
      editor.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
      filled.push("description");
    } catch (e) {
      console.error("[TourExt] fillQuill failed:", e.message);
      failed.push("description");
    }
  }

  async function runFill() {
    // Expand all collapsed mat-expansion-panels so Angular renders their content into the DOM.
    const collapsedHeaders = document.querySelectorAll(
      'mat-expansion-panel-header[aria-expanded="false"]'
    );
    collapsedHeaders.forEach((h) => h.click());
    if (collapsedHeaders.length > 0) await sleep(800);

    // Text / number inputs
    await fillByControl("tourTitle", tour.title, "title");
    await fillByControl("descriptionWillSee", tour.willSee, "willSee");
    await fillByControl("descriptionLearn", tour.willLearn, "willLearn");
    await fillByControl("mandatoryInformation", tour.mandatoryInfo, "mandatoryInfo");
    await fillByControl("recommendedInformation", tour.recommendedInfo, "recommendedInfo");
    await fillByControl("included", tour.included);
    await fillByControl("notIncluded", tour.notIncluded);
    await fillByControl("noOfPax", tour.noOfPax);
    await fillByControl("longitude", tour.longitude);
    await fillByControl("latitude", tour.latitude);
    await fillByControl("meetingPoint", tour.meetingPoint);
    await fillByControl("pickupInstructions", tour.pickupInstructions);
    await fillByControl("endPoint", tour.endPoint);
    await fillByControl("duration", tour.duration);
    // Prices
    await fillByControl("rate", tour.rate);
    await fillByControl("rateB2C", tour.rateB2C);
    await fillByControl("rate_request", tour.rateRequest, "rateRequest");
    await fillByControl("rate_requestB2C", tour.rateRequestB2C, "rateRequestB2C");
    await fillByControl("extraHourCharges", tour.extraHour, "extraHour");
    await fillByControl("extraHourChargesB2C", tour.extraHourB2C, "extraHourB2C");
    await fillByControl("extraHourCharges_request", tour.extraHourRequest, "extraHourRequest");
    await fillByControl("extraHourCharges_requestB2C", tour.extraHourRequestB2C, "extraHourRequestB2C");
    // Schedule
    await fillNgbDatepicker("startDate", tour.startDate);
    await fillNgbDatepicker("endDate", tour.endDate);
    await fillByControl("startTime", tour.startTime);
    await fillByControl("endTime", tour.endTime);
    // Cancellation & cut off — two separate fields for instant vs on-request
    await fillByControl("cancellation", tour.cancellation);
    await fillByControl("release", tour.release);

    // Quill
    fillQuill(tour.description);

    // ng-selects (sequential) — cascading: serviceType → activityType → subType
    await fillNgSelect("serviceType", tour.serviceType);
    await sleep(900); // wait for activityType options to cascade-load

    await fillNgSelect("activityType", tour.activityType);
    await sleep(900); // wait for subType options to cascade-load
    await fillNgSelect("subType", tour.subType);
    await fillNgSelect("activityFor", tour.activityFor);
    await fillNgSelect("voucherType", tour.voucherType);
    await fillNgSelect("countryId", tour.country, "country");
    await sleep(900); // wait for city options to cascade-load after country selection
    await fillNgSelect("cityId", tour.city, "city");

    await fillNgSelect("tourGuideLanguageList", tour.guideLanguageInstant, "guideLanguageInstant");
    await fillNgSelectMultiple("tourGuideLanguageList_request", tour.guideLanguageRequest, "guideLanguageRequest");
    await fillNgSelect("tagsList", tour.tags, "tags");

    return { success: true, filled, failed, errors };
  }

  return runFill();
})();
