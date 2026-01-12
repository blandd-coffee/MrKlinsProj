import express from "express";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import fs from "fs/promises";

const dataCache = {};

async function readJsonFile(filePath) {
  if (dataCache[filePath]) {
    return dataCache[filePath];
  }
  const data = await fs.readFile(filePath, "utf8");
  const parsed = JSON.parse(data);
  dataCache[filePath] = parsed;
  return parsed;
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
app.use(express.json());

// Define available data types
const AVAILABLE_TYPES = ["ECTS", "Mercer", "Corry", "Crawford"];
const DATA_DIR = join(__dirname, "../public");

app.get("/data", (req, res) => {
  if (!req.query.type || req.query.type === "") {
    // Return available data types when no type is specified
    res.json({ available: AVAILABLE_TYPES });
  } else {
    const type = req.query.type;

    // Validate type to prevent directory traversal
    if (!AVAILABLE_TYPES.includes(type)) {
      return res.status(400).json({
        error: `Invalid data type. Available types: ${AVAILABLE_TYPES.join(
          ", "
        )}`,
      });
    }

    const filename = `${type}.json`;
    const filepath = join(DATA_DIR, filename);

    readJsonFile(filepath)
      .then((data) => {
        res.json(data);
      })
      .catch((error) => {
        console.error(`Error reading file ${filepath}:`, error);
        res.status(500).json({
          error: "Failed to read data file",
          details: error.message,
        });
      });
  }
});

app.use(express.static(join(__dirname, "../public")));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
