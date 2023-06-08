function onInit(context) {
    // Pin the add-in's button to the ribbon.
    var ribbon = context.ui.ribbon;
    var tab = ribbon.tabs.getById("TabDefault");
    var group = tab.groups.getById("msgSendGroup");
    var control = group.controls.getById("msgSendOpenPaneButton");
    control.pinnable = true;
  }