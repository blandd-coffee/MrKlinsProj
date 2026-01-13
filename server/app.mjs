import express from "express";
import cors from "cors";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import fs from "fs/promises";

const dataCache = new Map();
const CACHE_TTL = 5 * 60 * 1000; // 5 minutes

async function readJsonFile(filePath) {
  const cached = dataCache.get(filePath);
  if (cached && Date.now() - cached.timestamp < CACHE_TTL) {
    return cached.data;
  }

  try {
    const data = await fs.readFile(filePath, "utf8");
    const parsed = JSON.parse(data);
    dataCache.set(filePath, { data: parsed, timestamp: Date.now() });
    return parsed;
  } catch (error) {
    if (error.code === "ENOENT") {
      throw new Error("Data file not found");
    }
    if (error instanceof SyntaxError) {
      throw new Error("Invalid JSON in data file");
    }
    throw error;
  }
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json());

// Define available data types
const AVAILABLE_TYPES = ["ECTS", "Mercer", "Corry", "Crawford"];
const DATA_DIR = join(__dirname, "../public");

app.get("/data", async (req, res) => {
  if (!req.query.type || req.query.type.trim() === "") {
    return res.json({ available: AVAILABLE_TYPES });
  }

  const type = req.query.type.trim();

  if (!AVAILABLE_TYPES.includes(type)) {
    return res.status(400).json({
      error: `Invalid data type. Available types: ${AVAILABLE_TYPES.join(
        ", "
      )}`,
    });
  }

  const filename = `${type}.json`;
  const filepath = join(DATA_DIR, filename);

  try {
    const data = await readJsonFile(filepath);
    res.json(data);
  } catch (error) {
    console.error(`Error reading file ${filepath}:`, error.message);
    const statusCode = error.message.includes("not found") ? 404 : 500;
    res.status(statusCode).json({ error: error.message });
  }
});

app.use(express.static(join(__dirname, "../public")));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
