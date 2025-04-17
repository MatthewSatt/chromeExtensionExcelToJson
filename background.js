// background.js
// This script runs in the background and helps manage the popup lifecycle

// Listen for messages from the popup
chrome.runtime.onMessage.addListener(function(message, sender, sendResponse) {
    // Handle the keepPopupOpen action
    if (message.action === "keepPopupOpen") {
        // We could implement additional logic here if needed
        // For now, just acknowledging the message keeps the connection alive
        sendResponse({status: "acknowledged"});

        // Return true to indicate we'll handle the response asynchronously
        return true;
    }
});

// Chrome extension lifetime events
chrome.runtime.onInstalled.addListener(function() {
    console.log("Excel to JSON extension installed");
});
