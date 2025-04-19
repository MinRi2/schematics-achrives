import { mkdir, readdir, readFile, rm, writeFile } from "fs/promises";
import path from "path";
import XLSX from "node-xlsx";
import { startPolling } from "./utils";

const OUT_PATH = path.resolve("./schematics");
const SCHEMATIC_SUFFIX = ".msch";

const QQ_DOC_COOKIES = Bun.env["QQ_DOC_COOKIES"]!;

const DOC_ID = "300000000$TshKyHrmMlQR";
const WISE_BOOK = "智能表1";

interface ExportData {
    ret: number;
    msg: string;
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
    let buffer: Buffer;

    if(Bun.env.NODE_ENV === "local"){
        buffer = await readFile(path.resolve("Mindustry 蓝图档案馆.xlsx"));
    }else{
        if (!QQ_DOC_COOKIES || QQ_DOC_COOKIES == "") {
            console.error("No QQ_DOC_COOKIES!");
            return;
        }

        try {
            const arrayBuffer = await fetchSchematicsExcel();
            buffer = Buffer.from(arrayBuffer);
        } catch (error) {
            console.error(error);
            console.error("Failed to fetch schematics excel. Please check your QQ_DOC_COOKIES.");
            return;
        }
    }

    await rm(OUT_PATH, {
        recursive: true,
        force: true
    });

    const data = await handleExcel(buffer);
    await genSchematics(data);
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

    let json: ExportData = await resp.json() as ExportData;
    if (json.ret != 0) {
        throw new Error(json.msg);
    }

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

    const schematicsData: SchematicData[] = excelData[0]!.data.map(arr => {
        const [category, author, name, _, base64] = arr;
        return { category, name, base64 };
    });

    return schematicsData;
}

async function genSchematics(schematicsData: SchematicData[]) {
    const jobs = schematicsData.map(async (data) => {
        const { category, name, base64 } = data;

        const handledName = name.replaceAll("/", "-").replaceAll("\n", "-");
        const fileName = handledName + SCHEMATIC_SUFFIX;
        const filePath = path.join(OUT_PATH, category, fileName);

        try {
            await mkdir(path.dirname(filePath), { recursive: true });
            await writeFile(filePath, Buffer.from(base64, "base64"), "binary");
            console.log("Save schematic", fileName);
        } catch (error) {
            console.error("Failed to save", filePath, error);
        }
    });

    await Promise.all(jobs);

    console.log("All schematics saved to", OUT_PATH);
    console.log("Total schematics:", schematicsData.length);
}

async function readData() {
    const schematicsData: SchematicData[] = [];

    const fileNames = await readdir(path.resolve(OUT_PATH));
    const readDataJobs = fileNames.map(async (category) => {
        console.log("Read category:", category);

        const categoryPath = path.join(OUT_PATH, category);
        const schematicFiles = await readdir(categoryPath);

        const readJobs = schematicFiles.map(async (schematicFileName) => {
            const schematicFilePath = path.join(categoryPath, schematicFileName);

            try {
                const string = await readFile(schematicFilePath, "binary");
                const base64 = Buffer.from(string, "binary").toBase64();
                schematicsData.push({ category, name: path.basename(schematicFileName, SCHEMATIC_SUFFIX), base64 });
                console.log("Read schematic", schematicFileName);
            } catch (error) {
                console.error("Failed to read schematic:", schematicFilePath, error);
            }
        });

        await Promise.all(readJobs);
    });

    await Promise.all(readDataJobs);

    return schematicsData;
}