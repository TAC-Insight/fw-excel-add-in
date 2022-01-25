export default async function setAlertMsg(msg) {
  // eslint-disable-next-line no-undef
  const alertMsg = document.getElementById("alertMsg");
  alertMsg.innerHTML = msg;
}
