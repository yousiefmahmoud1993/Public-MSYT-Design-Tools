// Saving files inside the Android/iOS app (Capacitor). On the web this file does nothing.
// The app calls window.nativeSave(blob, fileName) for PDF and Excel exports when running natively.
(function(){
  function isNative(){ return !!(window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform()); }
  window.isNativeApp = isNative;
  window.nativeSave = function(blob, name){
    var P = window.Capacitor.Plugins, r = new FileReader();
    r.onload = function(){
      var b64 = String(r.result).split(",")[1];
      P.Filesystem.writeFile({ path: name, data: b64, directory: "CACHE" })
        .then(function(res){ return P.Share.share({ title: name, files: [res.uri], dialogTitle: "Save or share" }); })
        .catch(function(e){ if (!/cancel/i.test(String(e && e.message))) alert("Export failed: " + (e && e.message || e)); });
    };
    r.readAsDataURL(blob);
  };
  window.nativePdf = window.nativeSave;
})();

// Phone Back (Android hardware/gesture back): go back inside the app (tool -> list -> all programs),
// and close the app only from the home screen. iOS uses the edge swipe (enabled in AppDelegate.swift).
(function(){
  if (!(window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform())) return;
  var App = window.Capacitor.Plugins.App;
  if (!App) return;
  App.addListener("backButton", function(ev){
    if (ev && ev.canGoBack) window.history.back();
    else App.exitApp();
  });
})();
