const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzd9K3q_yMCCNXcUm6DkZbhPRZ3GDhUil6J_1naX2FNqO1Aqbj70N0fYX17WO40JNpD/exec";

const form = document.getElementById("tshirtForm");
const submitBtn = document.getElementById("submitBtn");
const popup = document.getElementById("successPopup");
const popupClose = document.getElementById("popupClose");
const previewNumber = document.getElementById("previewNumber");
const previewName = document.getElementById("previewName");
const jerseyPreview = document.getElementById("jerseyPreview");

const fields = {
  fullName: document.getElementById("fullName"),
  phone: document.getElementById("phone"),
  size: document.getElementById("size"),
  sleeveLength: document.getElementById("sleeveLength"),
  lowerSize: document.getElementById("lowerSize"),
  backNumber: document.getElementById("backNumber"),
  backName: document.getElementById("backName"),
};

function setError(input, message) {
  const field = input.closest(".field");
  const errorEl = field.querySelector(".error");
  field.classList.toggle("invalid", Boolean(message));
  errorEl.textContent = message || "";
}

function validateName(value) {
  if (!value.trim()) return "Please enter your name.";
  if (value.trim().length < 2) return "Please enter a valid name.";
  return "";
}

function validatePhone(value) {
  if (!/^[6-9]\d{9}$/.test(value)) {
    return "Please enter a valid 10-digit mobile number.";
  }
  return "";
}

function validateSelect(value, label) {
  return value ? "" : `Please select a ${label}.`;
}

function validateBackNumber(value) {
  if (value === "") return "Please enter a jersey number.";
  if (!/^\d{1,3}$/.test(value)) {
    return "The jersey number must be 1 to 3 digits.";
  }
  return "";
}

function validateBackName(value) {
  if (!value.trim()) return "Please enter the name to print on the jersey.";
  if (!/^[A-Za-z ]+$/.test(value.trim())) {
    return "The jersey name may contain letters only.";
  }
  return "";
}

function validateForm() {
  const errors = {
    fullName: validateName(fields.fullName.value),
    phone: validatePhone(fields.phone.value),
    size: validateSelect(fields.size.value, "t-shirt size"),
    sleeveLength: validateSelect(fields.sleeveLength.value, "sleeve length"),
    lowerSize: validateSelect(fields.lowerSize.value, "trouser size"),
    backNumber: validateBackNumber(fields.backNumber.value),
    backName: validateBackName(fields.backName.value),
  };

  Object.keys(errors).forEach((key) => setError(fields[key], errors[key]));
  return !Object.values(errors).some(Boolean);
}

function updatePreview() {
  const number = fields.backNumber.value;
  const name = fields.backName.value.trim().toUpperCase();
  previewNumber.textContent = number === "" ? "10" : number;
  previewName.textContent = name || "NAME";
}

function showPopup() {
  popup.hidden = false;
  popupClose.focus();
}

function hidePopup() {
  popup.hidden = true;
}

function setLoading(isLoading) {
  submitBtn.disabled = isLoading;
  submitBtn.classList.toggle("loading", isLoading);
}

function getFormPayload() {
  syncSelectedSize();
  syncSelectedSleeveLength();
  syncSelectedLowerSize();
  return {
    fullName: fields.fullName.value.trim(),
    phone: fields.phone.value.trim(),
    size: fields.size.value,
    sleeveLength: fields.sleeveLength.value,
    lowerSize: fields.lowerSize.value,
    backNumber: String(fields.backNumber.value),
    backName: fields.backName.value.trim().toUpperCase(),
  };
}

async function saveToGoogleSheet(payload) {
  if (!GOOGLE_SCRIPT_URL || GOOGLE_SCRIPT_URL.includes("PASTE_YOUR_GOOGLE")) {
    console.warn("Google Apps Script URL is not set yet. Showing the success message locally.");
    return;
  }

  const body = new FormData();
  Object.entries(payload).forEach(([key, value]) => body.append(key, value));

  await fetch(GOOGLE_SCRIPT_URL, {
    method: "POST",
    mode: "no-cors",
    body,
  });
}

const sizeOptions = document.querySelectorAll('input[name="sizeOption"]');
const sleeveLengthOptions = document.querySelectorAll('input[name="sleeveLengthOption"]');
const lowerSizeOptions = document.querySelectorAll('input[name="lowerSizeOption"]');

function syncSelectedSize() {
  const selected = document.querySelector('input[name="sizeOption"]:checked');
  fields.size.value = selected ? selected.value : "";
  sizeOptions.forEach((item) => {
    item.closest(".size-card").classList.toggle("is-selected", item.checked);
  });
}

function syncSelectedSleeveLength() {
  const selected = document.querySelector('input[name="sleeveLengthOption"]:checked');
  fields.sleeveLength.value = selected ? selected.value : "";
  sleeveLengthOptions.forEach((item) => {
    item.closest(".size-card").classList.toggle("is-selected", item.checked);
  });
  jerseyPreview.classList.toggle("is-full-sleeve", fields.sleeveLength.value === "Full Sleeve");
}

function syncSelectedLowerSize() {
  const selected = document.querySelector('input[name="lowerSizeOption"]:checked');
  fields.lowerSize.value = selected ? selected.value : "";
  lowerSizeOptions.forEach((item) => {
    item.closest(".size-card").classList.toggle("is-selected", item.checked);
  });
}

sizeOptions.forEach((option) => {
  option.addEventListener("change", () => {
    syncSelectedSize();
    setError(fields.size, validateSelect(fields.size.value, "t-shirt size"));
  });
});

sleeveLengthOptions.forEach((option) => {
  option.addEventListener("change", () => {
    syncSelectedSleeveLength();
    setError(fields.sleeveLength, validateSelect(fields.sleeveLength.value, "sleeve length"));
  });
});

lowerSizeOptions.forEach((option) => {
  option.addEventListener("change", () => {
    syncSelectedLowerSize();
    setError(fields.lowerSize, validateSelect(fields.lowerSize.value, "trouser size"));
  });
});

fields.backNumber.addEventListener("input", () => {
  fields.backNumber.value = fields.backNumber.value.replace(/\D/g, "").slice(0, 3);
  updatePreview();
});

fields.backName.addEventListener("input", () => {
  fields.backName.value = fields.backName.value.toUpperCase();
  updatePreview();
});

fields.phone.addEventListener("input", () => {
  fields.phone.value = fields.phone.value.replace(/\D/g, "").slice(0, 10);
});

Object.values(fields).forEach((input) => {
  input.addEventListener("blur", () => {
    if (input === fields.fullName) setError(input, validateName(input.value));
    if (input === fields.phone) setError(input, validatePhone(input.value));
    if (input === fields.size) setError(input, validateSelect(input.value, "t-shirt size"));
    if (input === fields.sleeveLength) setError(input, validateSelect(input.value, "sleeve length"));
    if (input === fields.lowerSize) setError(input, validateSelect(input.value, "trouser size"));
    if (input === fields.backNumber) setError(input, validateBackNumber(input.value));
    if (input === fields.backName) setError(input, validateBackName(input.value));
  });
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  syncSelectedSize();
  syncSelectedSleeveLength();
  syncSelectedLowerSize();
  if (!validateForm()) return;

  setLoading(true);
  try {
    await saveToGoogleSheet(getFormPayload());
    form.reset();
    syncSelectedSize();
    syncSelectedSleeveLength();
    syncSelectedLowerSize();
    updatePreview();
    showPopup();
  } catch (error) {
    console.error(error);
    alert("Unable to submit your details. Please check your internet connection and try again.");
  } finally {
    setLoading(false);
  }
});

popupClose.addEventListener("click", hidePopup);
popup.addEventListener("click", (event) => {
  if (event.target === popup) hidePopup();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !popup.hidden) hidePopup();
});

updatePreview();
