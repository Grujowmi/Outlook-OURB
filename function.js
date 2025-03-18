// Initialize the Office Add-in.
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

function openTicketing(event) {
    var url = "https://link-you-want-to-acces.com/";
    Office.context.ui.openBrowserWindow(url);
    event.completed();
}
