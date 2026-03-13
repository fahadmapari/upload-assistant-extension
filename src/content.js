// Content script — runs on the admin panel page
// Listens for messages from the popup (optional direct messaging approach)

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === 'PING') {
    sendResponse({ alive: true });
  }
  return true;
});
