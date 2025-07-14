const express = require("express");
const multer = require("multer");
const path = require("path");
const { spawn } = require("child_process");
const fs = require("fs");

const app = express();
const port = 3000;

// Dossier pour stocker temporairement les uploads
const upload = multer({ dest: "uploads/" });

// Servir fichier statique (le frontend)
app.use(express.static("public"));

// Option : chemin complet vers python dans venv (décommente si besoin et adapte)
// const pythonExe = path.join(__dirname, "venv", "Scripts", "python.exe"); // Windows
// const pythonExe = path.join(__dirname, "venv", "bin", "python"); // Linux/macOS

// Route POST pour recevoir fichiers
// pdfs: array, excel: single
app.post(
  "/extract",
  upload.fields([
    { name: "pdfs", maxCount: 10 },
    { name: "excel", maxCount: 1 },
  ]),
  (req, res) => {
    try {
      if (
        !req.files["pdfs"] ||
        req.files["pdfs"].length === 0 ||
        !req.files["excel"] ||
        req.files["excel"].length === 0
      ) {
        return res.status(400).send("Fichiers manquants");
      }

      // Récupérer les chemins des fichiers uploadés
      const pdfPaths = req.files["pdfs"].map((file) => file.path);
      const excelPath = req.files["excel"][0].path;

      // Chemin pour le fichier Excel résultat
      const outputExcelPath = path.join("uploads", "resultat.xlsx");

      // Appel du script Python en passant les chemins des fichiers en argument
      // Remplace "python" par pythonExe si tu veux forcer le python du venv
      const pythonProcess = spawn(
        "python",
        // pythonExe,
        [
          "traitement.py",
          pdfPaths.join(","),
          excelPath,
          outputExcelPath,
        ]
      );

      pythonProcess.stdout.on("data", (data) => {
        console.log(`Python stdout: ${data}`);
      });

      pythonProcess.stderr.on("data", (data) => {
        console.error(`Python stderr: ${data}`);
      });

      pythonProcess.on("close", (code) => {
        console.log(`Python process exited with code ${code}`);
        if (code === 0) {
          // Envoyer le fichier Excel résultat au client
          res.download(outputExcelPath, "matrice_indicateurs_mise_a_jour.xlsx", (err) => {
            if (err) {
              console.error("Erreur lors de l'envoi du fichier:", err);
              res.status(500).send("Erreur lors de l'envoi du fichier");
            }

            // Nettoyer fichiers uploadés
            [...pdfPaths, excelPath, outputExcelPath].forEach((file) => {
              fs.unlink(file, () => {});
            });
          });
        } else {
          res.status(500).send("Erreur dans le traitement Python");
        }
      });
    } catch (e) {
      console.error(e);
      res.status(500).send("Erreur serveur");
    }
  }
);

app.listen(port, () => {
  console.log(`Serveur démarré sur http://localhost:${port}`);
});
