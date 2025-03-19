Office.onReady(function() {
    
});

function openTicketingTool(event) {
    var url = "https://target-url.com";
    Office.context.ui.openBrowserWindow(url);
    event.completed();
}

Office.actions.associate("openTicketingTool", openTicketingTool);
