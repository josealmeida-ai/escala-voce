// Publica o site (Firebase Hosting) a partir da raiz do projeto.
// Uso: npm run deploy
const { execSync } = require("child_process");
const readline = require("readline");

function run(cmd) {
  execSync(cmd, { stdio: "inherit" });
}

function ask(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => rl.question(question, (answer) => {
    rl.close();
    resolve(answer.trim().toLowerCase());
  }));
}

(async () => {
  try {
    const status = execSync("git status --porcelain", { encoding: "utf8" });
    if (status.trim()) {
      console.log("Há alterações não commitadas:\n" + status);
      const answer = await ask("Continuar o deploy mesmo assim? (s/N) ");
      if (answer !== "s" && answer !== "sim") {
        console.log("Deploy cancelado.");
        process.exit(1);
      }
    }

    console.log("Publicando no Firebase Hosting (projeto: escala-voce)...");
    run("firebase deploy --only hosting");
    console.log("Deploy concluído.");
  } catch (err) {
    console.error("Falha no deploy:", err.message);
    process.exit(1);
  }
})();
