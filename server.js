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
    { name: "document", maxCount: 10 },
    { name: "excel", maxCount: 1 },
  ]),
  (req, res) => {
    try {
      const docFiles = req.files["document"];
      const excelFile = req.files["excel"]?.[0];

      if (!docFiles || docFiles.length === 0 || !excelFile) {
        return res.status(400).send("Fichiers manquants");
      }

      // Vérifie que tous les documents sont du même type
      const extensions = docFiles.map(f => path.extname(f.originalname).toLowerCase());
      const allArePDF = extensions.every(ext => ext === ".pdf");
      const allAreWord = extensions.every(ext => ext === ".doc" || ext === ".docx");

      if (!allArePDF && !allAreWord) {
        return res.status(400).send("Tous les fichiers doivent être soit PDF, soit Word");
      }

      const docType = allArePDF ? "pdf" : "word";
      const docPaths = docFiles.map(f => f.path);
      const docPathsCSV = docPaths.join(",");
      const outputExcelPath = path.join("uploads", "resultat.xlsx");

      const pythonProcess = spawn("python", [
        "traitement.py",
        docType,
        docPathsCSV,
        excelFile.path,
        outputExcelPath,
      ]);

      pythonProcess.stdout.on("data", (data) => {
        console.log(`Python stdout: ${data}`);
      });

      pythonProcess.stderr.on("data", (data) => {
        console.error(`Python stderr: ${data}`);
      });

      pythonProcess.on("close", (code) => {
        console.log(`Python process exited with code ${code}`);
        if (code === 0) {
          res.download(outputExcelPath, "matrice_indicateurs_mise_a_jour.xlsx", (err) => {
            if (err) {
              console.error("Erreur lors de l'envoi du fichier:", err);
              res.status(500).send("Erreur lors de l'envoi du fichier");
            }

            // Nettoyage fichiers temporaires
            [...docPaths, excelFile.path, outputExcelPath].forEach((file) => {
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