/* global document, Office, Word */

const targetWords = ["Word", "Office", "script"];

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("target-words").innerText = targetWords.join(", ");
    document.getElementById("insert").onclick = insertExerciseText;
    document.getElementById("check").onclick = checkWords;
  }
});

export async function insertExerciseText() {
  return Word.run(async (context) => {
    context.document.body.clear();
    const longText = `Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.` +
      `Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.` +
      `Lorem ipsum dolor sit amet, consectetur Script elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.` +
      `Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.` +
      `Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.` +
      `Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco Office laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum. Ici nous mentionnons Word ` +
      `Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.`

    context.document.body.insertParagraph(longText, Word.InsertLocation.start);
    await context.sync();
  });
}

export async function checkWords() {
  return Word.run(async (context) => {
    const targetSet = new Set(targetWords.map(w => w.toLowerCase()));
    const bodyRange = context.document.body.getRange();
    const delimiters = [" ", "\n", "\r", "\t", ".", ",", ";", ":", "!", "?", "(", ")", "\"", "«", "»", "—", "-", "“", "”", "’", "'", "…"];
    const textRanges = bodyRange.getTextRanges(delimiters, true);
    context.load(textRanges, "text,font/highlightColor");
    await context.sync();

    const isYellow = h => {
      if (!h) return false;
      const s = String(h).toLowerCase();
      return s.includes("yellow") || s === "#ffff00" || s === "yellow";
    };

    let allCorrect = true;
    const bad = { missing: new Set(), extra: new Set() };

    for (const r of textRanges.items) {
      let token = (r.text || "").trim();
      if (!token) continue;
      token = token.replace(/^[^A-Za-z0-9À-ÖØ-öø-ÿ]+|[^A-Za-z0-9À-ÖØ-öø-ÿ]+$/g, "");
      if (!token) continue;

      const lower = token.toLowerCase();
      const highlight = (r.font && r.font.highlightColor) ? r.font.highlightColor : "None";
      const yellow = isYellow(highlight);

      if (yellow && !targetSet.has(lower)) bad.extra.add(token);
      if (!yellow && targetSet.has(lower)) bad.missing.add(token);
    }

    if (bad.extra.size || bad.missing.size) allCorrect = false;

    const resultDiv = document.getElementById("result");
    if (allCorrect) {
      resultDiv.textContent = "✅ Exercice réussi : uniquement les bons mots sont surlignés en jaune";
      resultDiv.style.color = "green";
    } else {
      resultDiv.textContent = `❌ Incorrect. Manquants: ${[...bad.missing].join(", ") || "—"}; En trop: ${[...bad.extra].join(", ") || "—"}`;
      resultDiv.style.color = "red";
    }

    await context.sync();
  });
}

if (typeof window !== "undefined") {
  window.insertExerciseText = insertExerciseText;
  window.checkWords = checkWords;
}
