/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // affiche la liste des mots dans l'UI
    document.getElementById("target-words").innerText = targetWords.join(", ");

    // attache les boutons
    document.getElementById("insert").onclick = insertExerciseText;
    document.getElementById("check").onclick = checkWords;
  }
});

// liste des mots à souligner
const targetWords = ["Word", "Office", "script"];

// === Bouton "Insérer le texte d'exercice" ===
export async function insertExerciseText() {
  return Word.run(async (context) => {
    // efface tout le document
    context.document.body.clear();

    // texte long (2 pages environ selon police/mise en page)
    const longText = `Ceci est un texte d'exercice qui contient plusieurs paragraphes. Tu dois souligner certains mots précis comme Word, Office et script. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec a diam lectus. Sed sit amet ipsum mauris. Maecenas congue ligula ac quam viverra nec consectetur ante hendrerit. Donec et mollis dolor. Praesent et diam eget libero egestas mattis sit amet vitae augue. Nam tincidunt congue enim, ut porta lorem lacinia consectetur. Donec ut libero sed arcu vehicula ultricies a non tortor. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aenean ut gravida lorem. Ut turpis felis, pulvinar a semper sed, adipiscing id dolor. Pellentesque auctor nisi id magna consequat sagittis. Curabitur dapibus enim sit amet elit pharetra tincidunt feugiat nisl imperdiet. Vestibulum auctor dapibus neque. Nunc dignissim risus id metus. Cras ornare tristique elit. Vivamus vestibulum ntulla nec ante. Praesent placerat risus quis eros. Fusce pellentesque suscipit nibh. Integer vitae libero ac risus egestas placerat. Vestibulum commodo felis quis tortor. Ut aliquam sollicitudin leo. Cras iaculis ultricies nulla. Donec quis dui at dolor tempor interdum. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Donec a diam lectus. Sed sit amet ipsum mauris. Maecenas congue ligula ac quam viverra nec consectetur ante hendrerit. Donec et mollis dolor. Praesent et diam eget libero egestas mattis sit amet vitae augue. Nam tincidunt congue enim, ut porta lorem lacinia consectetur. Donec ut libero sed arcu vehicula ultricies a non tortor. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aenean ut gravida lorem. Ut turpis felis, pulvinar a semper sed, adipiscing id dolor. Pellentesque auctor nisi id magna consequat sagittis. Curabitur dapibus enim sit amet elit pharetra tincidunt feugiat nisl imperdiet. (tu peux dupliquer ces paragraphes autant de fois que nécessaire pour que le document fasse environ 2 pages complètes).`;

    // insertion
    context.document.body.insertParagraph(longText, Word.InsertLocation.start);

    await context.sync();
  });
}

// === Bouton "Contrôler les mots soulignés" ===
export async function checkWords() {
  return Word.run(async (context) => {
    let allUnderlined = true;

    for (let word of targetWords) {
      const searchResults = context.document.body.search(word, {
        matchCase: false,
        matchWholeWord: true,
      });

      context.load(searchResults, "font/underline");
      await context.sync();

      if (
        searchResults.items.length === 0 ||
        searchResults.items.some((item) => item.font.underline === "None")
      ) {
        allUnderlined = false;
        break;
      }
    }

    // affiche le résultat dans l’UI
    const resultDiv = document.getElementById("result");
    if (allUnderlined) {
      resultDiv.textContent = "✅ Tous les mots sont bien soulignés";
      resultDiv.style.color = "green";
    } else {
      resultDiv.textContent = "❌ Tous les mots ne sont pas soulignés";
      resultDiv.style.color = "red";
    }

    await context.sync();
  });
}

// expose en global si pas de bundler
if (typeof window !== "undefined") {
  window.insertExerciseText = insertExerciseText;
  window.checkWords = checkWords;
}
