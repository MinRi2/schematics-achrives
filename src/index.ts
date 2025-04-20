import { readdir, readFile, rm } from "fs/promises";
import path from "path";
import XLSX from "node-xlsx";
import { startPolling } from "./utils";
import { existsSync } from "fs";

import blacklist from "../blacklist.json";

class SchematicDataMap extends Map<string, SchematicData> {
    public static getKey(category: string, name: string) {
        return category + name;
    }

    public static getKeyMeta(meta: SchematicMeta) {
        return SchematicDataMap.getKey(meta.category, meta.name);
    }

    setData(param: SchematicData | SchematicData[]){
        if(param instanceof Array){
            if(param.length) param.forEach(meta => this.set(SchematicDataMap.getKeyMeta(meta), meta));
        }else if(param instanceof SchematicData){
            this.set(SchematicDataMap.getKeyMeta(param), param);
        }
    }

    getData(category: string, name: string) {
        return this.get(SchematicDataMap.getKey(category, name));
    }

    hasData(category: string, name: string) {
        return this.has(SchematicDataMap.getKey(category, name));
    }

    deleteData(data: SchematicMeta) {
        return this.delete(SchematicDataMap.getKey(data.category, data.name));
    }
}

class SchematicData implements SchematicMeta {
    public static SCHEMATIC_SUFFIX = ".msch";

    constructor(public category: string, public name: string, public base64: string) {
        this.name = name.replaceAll("/", "-").replaceAll("\n", "-");
    }

    getFileName() {
        return this.name + SchematicData.SCHEMATIC_SUFFIX;
    }

    getFilePath(basePath: string) {
        return path.resolve(basePath, this.category, this.getFileName());
    }

    equals(other: SchematicMeta) { 
        return this.category === other.category && this.name === other.name && this.base64 === other.base64;
    }
}

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

interface ContextData{
    qqCookies: string;
    outPath: string;
    lastData: SchematicDataMap;
    data: SchematicDataMap;
}

interface SchematicMeta{
    category: string;
    name: string;
    base64: string;
}

const context: ContextData = {
    outPath: "./schematics",
    qqCookies: Bun.env["QQ_DOC_COOKIES"]!,
    lastData: new SchematicDataMap(),
    data: new SchematicDataMap(),
}

const DOC_ID = "300000000$TshKyHrmMlQR";
const WISE_BOOK = "智能表1";

run();

async function run() {
    context.outPath = path.resolve(context.outPath);

    const { outPath, lastData, data } = context;
    lastData.setData(await readData(context.outPath));

    const excelData = await readExcel();
    if(excelData === undefined){
        console.error("Failed to parse excel data.");
        return;
    }

    data.setData(await parseExcelData(excelData[0]!.data));

    await genSchematics(outPath, data, lastData);
    await cleanSchematics(outPath, data, lastData);
}

async function fetchSchematicsExcel() {
    let resp = await fetch("https://docs.qq.com/v1/export/export_office", {
        "headers": {
            "content-type": "application/x-www-form-urlencoded;charset=UTF-8",
            "cookie": context.qqCookies,
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
                "cookie": context.qqCookies,
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

async function readExcel() {
    let buffer: Buffer;
    if(Bun.env.NODE_ENV === "local"){
        buffer = await readFile(path.resolve("Mindustry 蓝图档案馆.xlsx"));
    }else{
        if (!context.qqCookies || context.qqCookies == "") {
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

    const excelData = XLSX.parse(buffer, {
        type: "buffer",
        sheets: WISE_BOOK,
        cellHTML: false,
    });

    return excelData;
}

async function parseExcelData(excelData: any[]) {
    const schematicsData: SchematicData[] = excelData.filter((arr, index) => {
        if(index === 0) return false; // Skip header row

        const [category, author, name, _, base64] = arr;

        // schematics causing buffer mismatch will be listed.
        if(blacklist.schematics.findIndex(s => s === name) !== -1 
        || blacklist.authors.findIndex(a => a === author) !== -1){
            return false;
        }

        return isValidBase64(base64) || console.error("Invalid schematic", name);
    }).map(arr => {
        const [category, author, name, _, base64] = arr;
        return new SchematicData(category, name, base64);
    });

    return schematicsData;

    function isValidBase64(str: string) {
        return /^[A-Za-z0-9+/]+={0,2}$/.test(str) && (str.length % 4 === 0);
    }
}

async function genSchematics(fromPath: string, dataMap: SchematicDataMap, lastDataMap: SchematicDataMap) {
    let count = 0;
    const jobs = Array.from(dataMap.values()).map(async (data) => {
        if(lastDataMap.getData(data.category, data.name)?.equals(data)){
            return;
        }
        const { base64 } = data;
        const fileName = data.getFileName();
        const filePath = data.getFilePath(fromPath);
        
        const buffer = Buffer.from(base64, "base64");
        const file = Bun.file(filePath);

        await Bun.write(file, buffer);

        count++;
        console.log("Save schematic", fileName);
    });

    await Promise.all(jobs);

    console.log("All schematics saved to", context.outPath);
    console.log("Total schematics:", dataMap.size);
    console.log("Saved schematics:", count);
}

async function cleanSchematics(fromPath: string, dataMap: SchematicDataMap, lastDataMap: SchematicDataMap) {
    const jobs = Array.from(lastDataMap.entries()).map(async entry => {
        const [key, item] = entry;

        if(dataMap.has(key)) {
            return;
        }

        const filePath = item.getFilePath(fromPath);

        try {
            await rm(filePath, { force: true, recursive: true });
            console.log("Remove schematic", filePath);
        } catch (error) {
            console.error("Failed to remove schematic", filePath, error);
        }
    });

    await Promise.all(jobs);
}

async function readData(fromPath: string) {
    const schematicsData: SchematicData[] = [];

    if(!existsSync(fromPath)){
        return [];
    }

    const fileNames = await readdir(fromPath);
    const readDataJobs = fileNames.map(async (category) => {
        const categoryPath = path.join(fromPath, category);
        const schematicFiles = await readdir(categoryPath);

        const readJobs = schematicFiles.map(async (schematicFileName) => {
            const schematicFilePath = path.join(categoryPath, schematicFileName);

            try {
                const buffer = Buffer.from(await Bun.file(schematicFilePath).arrayBuffer());
                const base64 = buffer.toBase64();
                
                const data = new SchematicData(category, path.basename(schematicFileName, SchematicData.SCHEMATIC_SUFFIX), base64)
                schematicsData.push(data);
            } catch (error) {
                console.error("Failed to read schematic:", schematicFilePath, error);
            }
        });

        await Promise.all(readJobs);
    });

    await Promise.all(readDataJobs);

    console.log("Read schematics data:", schematicsData.length);

    return schematicsData;
}