// WPS loads main.js through its generated index.html. Keep the add-in itself tiny.
function LiveDocumentConnect() {
  var base='http://127.0.0.1:18765'; // Match LIVEDOC_PORT; for TLS use https.
  var pane=Application.CreateTaskPane(base+'/bridge/wps.html','Live Document');
  if(!pane) { Application.alert('Live Document: start the local runtime first (livedoc start).');return; }
  pane.Visible=true;
}
