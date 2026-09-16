// app.js
//
// Two directions of communication with Python (via pywebview):
//   JS -> Python:  pywebview.api.<method>(...)   (defined on the Api class in main.py)
//   Python -> JS:  window.evaluate_js("someFunctionDefinedHere(...)")
//
// Every function below that Python calls by name (addMessage, setStatus,
// pulseHUD, showConfirm, setMicState) is intentionally global (attached to
// window) so main.py's evaluate_js() calls can find it.

const chatLog = document.getElementById("chat-log");
const hudWrap = document.getElementById("hud-wrap");
const statusEl = document.getElementById("status");
const inputBar = document.getElementById("input-bar");
const textInput = document.getElementById("text-input");
const sendBtn = document.getElementById("send-btn");
const micBtn = document.getElementById("mic-btn");
const confirmOverlay = document.getElementById("confirm-overlay");
const confirmText = document.getElementById("confirm-text");
const confirmYes = document.getElementById("confirm-yes");
const confirmNo = document.getElementById("confirm-no");
const apikeyOverlay = document.getElementById("apikey-overlay");
const apikeyInput = document.getElementById("apikey-input");
const apikeySave = document.getElementById("apikey-save");

hudWrap.classList.add("idle");

// ---------------- click feedback ----------------

// Retriggerable bounce animation for any element - removes then re-adds
// the class so rapid repeated clicks each replay it instead of only the
// first one animating.
function popAnimate(el) {
  el.classList.remove("btn-pop");
  void el.offsetWidth; // force reflow so the removed class actually registers
  el.classList.add("btn-pop");
}
[sendBtn, micBtn, confirmYes, confirmNo, apikeySave].forEach((el) => {
  el.addEventListener("click", () => popAnimate(el));
});
hudWrap.addEventListener("click", () => popAnimate(hudWrap));

// ---------------- functions Python calls into ----------------

window.addMessage = function (sender, text) {
  const div = document.createElement("div");
  div.className = "msg " + (sender === "user" ? "user" : sender === "action" ? "action" : "jarvis");
  div.textContent = text;
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
};

window.setStatus = function (status) {
  // status: "STANDBY" | "LISTENING" | "PROCESSING" | "SPEAKING"
  statusEl.textContent = status;
  hudWrap.classList.remove("idle", "listening");
  if (status === "LISTENING") {
    hudWrap.classList.add("listening");
  } else if (status === "STANDBY") {
    hudWrap.classList.add("idle");
  }
};

let pulseTimeout = null;
window.pulseHUD = function () {
  hudWrap.classList.remove("idle");
  hudWrap.classList.add("pulse");
  clearTimeout(pulseTimeout);
  pulseTimeout = setTimeout(() => hudWrap.classList.remove("pulse"), 140);
};

let micListening = false;
window.setMicState = function (state) {
  // state: "hidden" | "idle" | "recording"
  micListening = state === "recording";
  micBtn.classList.remove("recording", "hidden");
  if (state === "hidden") micBtn.classList.add("hidden");
  if (state === "recording") micBtn.classList.add("recording");
};

window.setInputEnabled = function (enabled) {
  textInput.disabled = !enabled;
  sendBtn.disabled = !enabled;
  micBtn.disabled = !enabled;
};

window.showConfirm = function (text) {
  confirmText.textContent = text;
  confirmOverlay.classList.remove("hidden");
};

window.showApiKeyPrompt = function () {
  apikeyOverlay.classList.remove("hidden");
  apikeyInput.focus();
};

function hideConfirm() {
  confirmOverlay.classList.add("hidden");
}

// ---------------- JS -> Python ----------------

function sendCurrentMessage() {
  const text = textInput.value.trim();
  if (!text) return;
  window.addMessage("user", text);
  textInput.value = "";
  window.pywebview.api.send_message(text);
  textInput.blur();
  inputBar.classList.remove("expanded");
}

sendBtn.addEventListener("click", sendCurrentMessage);
textInput.addEventListener("keydown", (e) => {
  if (e.key === "Enter") sendCurrentMessage();
});
textInput.addEventListener("focus", () => {
  inputBar.classList.add("expanded");
});

micBtn.addEventListener("click", () => {
  // Flip the button's look immediately rather than waiting on a Python
  // round-trip (which, on some WebView2 setups, can lag or occasionally
  // drop) - this is purely visual optimism; toggle_mic() below is what
  // actually starts/stops the recording, and setMicState() will correct
  // the button if Python's real state ever disagrees with this guess.
  micListening = !micListening;
  micBtn.classList.toggle("recording", micListening);

  window.pywebview.api.toggle_mic().then((ok) => {
    if (!ok) {
      // Rejected (e.g. mic unavailable) - undo the optimistic flip.
      micListening = false;
      micBtn.classList.remove("recording");
    }
  });
});

hudWrap.addEventListener("click", () => {
  // Tap the HUD orb to stop JARVIS talking without starting the mic -
  // a silent "quiet" button distinct from the mic's barge-in-and-listen.
  window.pywebview.api.interrupt();
});

confirmYes.addEventListener("click", () => {
  hideConfirm();
  window.pywebview.api.confirm_response(true);
});
confirmNo.addEventListener("click", () => {
  hideConfirm();
  window.pywebview.api.confirm_response(false);
});

function saveApiKey() {
  const key = apikeyInput.value.trim();
  if (!key) return;
  apikeyOverlay.classList.add("hidden");
  window.pywebview.api.save_api_key(key);
}
apikeySave.addEventListener("click", saveApiKey);
apikeyInput.addEventListener("keydown", (e) => {
  if (e.key === "Enter") saveApiKey();
});
