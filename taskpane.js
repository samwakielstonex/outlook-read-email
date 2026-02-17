Office.onReady(() => {
  const btn = document.getElementById("readBodyBtn");
  if (btn) btn.addEventListener("click", readBody);
});

function readBody() {
  const item = Office.context.mailbox?.item;
  if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
    write("Please select an email first.");
    return;
  }

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      write(result.value || "");
    } else {
      write("Error: " + result.error.message);
    }
  });
}

function write(text) {
  document.getElementById("output").textContent = text || "";
}