import { mkdir, readdir, readFile, writeFile } from "fs/promises";
import path from "path";
import XLSX from "node-xlsx";
import { startPolling } from "./utils";

const OUT_PATH = "./schematics";
const SCHEMATIC_SUFFIX = ".mesh";

const QQ_DOC_COOKIES = process.env["QQ_DOC_COOKIES"]!;

const DOC_ID = "300000000$TshKyHrmMlQR";
const WISE_BOOK = "智能表1";

interface ExportData {
    operationId: string;
}

interface QueryData {
    status: "Processing" | "Done";
    progress: number,
    file_url?: string;
    file_name?: string;
    file_size?: number;
}

interface SchematicData {
    category: string;
    name: string;
    base64: string;
}

run();

async function run() {
    if (!QQ_DOC_COOKIES || QQ_DOC_COOKIES == "") {
        console.error("No QQ_DOC_COOKIES!");
        return;
    }

    const arrayBuffer = await fetchSchematicsExcel();
    const buffer = Buffer.from(arrayBuffer);

    // read locally
    // const buffer = await readFile(path.resolve("Mindustry 蓝图档案馆.xlsx"));
    await handleExcel(buffer);
}

async function fetchSchematicsExcel() {
    let resp = await fetch("https://docs.qq.com/v1/export/export_office", {
        "headers": {
            "content-type": "application/x-www-form-urlencoded;charset=UTF-8",
            "cookie": QQ_DOC_COOKIES,
        },
        "body": `exportType=0&switches=%7B%22embedFonts%22%3Afalse%7D&docId=${DOC_ID}`,
        "method": "POST",
    });

    let json: ExportData = await resp.json();

    let operationId = json.operationId;
    const data = await startPolling(async () => {
        const resp = await fetch(`https://docs.qq.com/v1/export/query_progress?operationId=${operationId}`, {
            "headers": {
                "cookie": QQ_DOC_COOKIES,
            },
            "method": "GET"
        });

        return await resp.json() as QueryData;
    }, {
        finished(result) {
            return result.status == "Done";
        },
    });

    console.log("Fetch", data.file_name);

    resp = await fetch(data.file_url!, {
        method: "GET",
    });

    return await resp.arrayBuffer();
}

async function handleExcel(buffer: Buffer) {
    const excelData = XLSX.parse(buffer, {
        type: "buffer",
        sheets: WISE_BOOK,
        cellHTML: false,
    });

    const schematicsData: SchematicData[] = excelData[0].data.map(arr => {
        const [category, author, name, _, base64] = arr;
        return { category, name, base64 };
    });

    const jobs = schematicsData.map(async (data) => {
        const { category, name, base64 } = data;

        const handledName = name.replaceAll("/", "-").replaceAll("\n", "-");
        const fileName = handledName + SCHEMATIC_SUFFIX;
        const filePath = path.resolve(OUT_PATH, category, fileName);

        try {
            await mkdir(path.dirname(filePath), { recursive: true });
            await writeFile(filePath, Buffer.from(base64, "base64"));
        } catch (error) {
            console.error("Failed to save", filePath, error);
        }
    });

    await Promise.all(jobs);
}

async function readData() {
    const schematicsData: SchematicData[] = [];

    const fileNames = await readdir(path.resolve(OUT_PATH));
    const readDataJobs = fileNames.map(async (category) => {
        console.log("Read category:", category);

        const categoryPath = path.resolve(OUT_PATH, category);
        const schematicFiles = await readdir(categoryPath);

        const readJobs = schematicFiles.map(async (schematicFileName) => {
            const schematicFilePath = path.resolve(categoryPath, schematicFileName);

            try {
                const string = await readFile(schematicFilePath, "utf-8");
                const base64 = Buffer.from(string, "utf-8").toString("base64");
                schematicsData.push({ category, name: path.basename(schematicFileName), base64 });
            } catch (error) {
                console.error("Failed to read schematic:", schematicFilePath, error);
            }
        });

        await Promise.all(readJobs);
    });

    await Promise.all(readDataJobs);

    return schematicsData;
}