/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("app").textContent = "SQL Query Taskpane";
  }
});
