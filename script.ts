import { MendixPlatformClient, OnlineWorkingCopy } from "mendixplatformsdk";
import { ModelSdkClient, IModel, projects, domainmodels, microflows, pages, navigation, texts, security, IStructure, menus, IList, nanoflows } from "mendixmodelsdk";
import * as fs from 'fs';
import * as stream from 'stream';
import * as officegen from 'officegen';
import * as ExcelJS from 'exceljs';

const v8 = require('v8');
v8.setFlagsFromString('--expose_gc');
const gc = global.gc;

// const appId = "{{AppID}}";
const appId = "c99be9a4-8ccf-4c29-aabb-7ea0c7242ebc";
const branchName = null // null for mainline
const wc = null;
const client = new MendixPlatformClient();
let pObj;

const CHUNK_SIZE = 10000; // Verder verhoogd aantal items om per keer te verwerken

// Functie om lange strings af te kappen
function truncateString(str: string, maxLength: number = 32000): string {
    if (str && str.length > maxLength) {
        return str.substring(0, maxLength) + '...';
    }
    return str;
}

// Functie om een nieuwe Excel workbook en worksheets te maken
function createWorkbook() {
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet('Entities').addRow(['User Role', 'Module', 'Module Role', 'Entity', 'Xpath', 'Create/Delete', 'Member Rules']);
    workbook.addWorksheet('Pages').addRow(['User Role', 'Module', 'Module Role', 'Page Name', 'Allowed']);
    workbook.addWorksheet('Microflows').addRow(['User Role', 'Module', 'Module Role', 'Microflows', 'Allowed']);
    workbook.addWorksheet('Nanoflows').addRow(['User Role', 'Module', 'Module Role', 'Nanoflows', 'Allowed']);
    return workbook;
}

function getWorksheet(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet {
    return workbook.getWorksheet(name);
}

/*
 * PROJECT TO ANALYZE
 */
const app = client.getApp(appId);
main();

process.on('unhandledRejection', (reason, promise) => {
    console.log('Unhandled Rejection at:', reason.stack || reason)
});

process.on('warning', (warning) => {
    console.warn(warning.name);    // Print the warning name
    console.warn(warning.message); // Print the warning message
    console.warn(warning.stack);   // Print the stack trace
});

async function main() {
    const repository = app.getRepository();
    const useBranch = branchName === null
        ? (await repository.getInfo()).type === 'svn' ? 'trunk' : 'main'
        : branchName;

    const workingCopy = await app.createTemporaryWorkingCopy(useBranch);
    const projectSecurity = await loadProjectSecurity(workingCopy);
    const userRoles = getAllUserRoles(projectSecurity);

    const workbook = createWorkbook();
    await createUserSecurityDocument(userRoles, workbook);

    const writeStream = fs.createWriteStream('MendixSecurityDocument.xlsx');
    await workbook.xlsx.write(writeStream);
    writeStream.on('finish', () => {
        console.log('Finished creating Document');
        // Forceer garbage collection
        if (global.gc) {
            global.gc();
        }
    });
}

/**
* This function picks the first navigation document in the project.
*/
async function createUserSecurityDocument(userRoles: security.UserRole[], workbook: ExcelJS.Workbook) {
    console.log("Creating User Access Matrix");
    for (let i = 0; i < userRoles.length; i += CHUNK_SIZE) {
        const chunk = userRoles.slice(i, i + CHUNK_SIZE);
        await Promise.all(chunk.map(async (userRole) => processAllModules(userRole, workbook)));
        // Forceer garbage collection na elke chunk
        if (global.gc) {
            global.gc();
        }
    }
}

async function processAllModules(userRole: security.UserRole, workbook: ExcelJS.Workbook): Promise<void> {
    const modules = userRole.model.allModules();
    for (let i = 0; i < modules.length; i += CHUNK_SIZE) {
        const chunk = modules.slice(i, i + CHUNK_SIZE);
        await Promise.all(chunk.map(async (module) => processModule(module, userRole, workbook)));
        // Forceer garbage collection na elke chunk
        if (global.gc) {
            global.gc();
        }
    }
}

async function processModule(module: projects.IModule, userRole: security.UserRole, workbook: ExcelJS.Workbook): Promise<void> {
    // console.debug(`Processing module: ${module.name}`);
    var securities = await getAllModuleSecurities(module);
    await Promise.all(securities.map(async (security) => loadAllModuleSecurities(securities, userRole, workbook)));
}
async function getAllModuleSecurities(module: projects.IModule): Promise<security.IModuleSecurity[]> {
    // console.debug(`Processing getAllModuleSecurities: ${module.name}`);
    return module.model.allModuleSecurities().filter(modSecurity => {
        if (modSecurity != null) {
            console.debug(`Mod Security is not null: ${modSecurity.containerAsModule.name}`);
            return modSecurity.containerAsModule.name === module.name;
        } else {
            return false;
        };

    });
}

async function loadAllModuleSecurities(moduleSecurities: security.IModuleSecurity[], userRole: security.UserRole, workbook: ExcelJS.Workbook): Promise<void> {
    await Promise.all(moduleSecurities.map(async (mSecurity) => processLoadedModSec(mSecurity, userRole, workbook)));
}

async function processLoadedModSec(modSec: security.IModuleSecurity, userRole: security.UserRole, workbook: ExcelJS.Workbook): Promise<void> {
    await Promise.all(modSec.moduleRoles.map(async (modRole) => processModRole(modRole, userRole, workbook)));
}



async function loadModSec(modSec: security.IModuleSecurity): Promise<security.ModuleSecurity> {
    // console.debug(`Processing loadModSec`);
    return modSec.load();
}



async function processModRole(modRole: security.IModuleRole, userRole: security.UserRole, workbook: ExcelJS.Workbook): Promise<void> {
    if (addIfModuleRoleInUserRole(modRole, userRole)) {
        await Promise.all(modRole.containerAsModuleSecurity.containerAsModule.domainModel.entities.map(async (entity) =>
            processAllEntitySecurityRules(entity, modRole, userRole, workbook)
                .then(() => processAllPages(modRole, userRole, workbook))
                .then(() => processAllMicroflows(modRole, userRole, workbook))
                .then(() => processAllNanoflows(modRole, userRole, workbook))));
    }
}
async function processAllEntitySecurityRules(entity: domainmodels.IEntity, moduleRole: security.IModuleRole, userRole: security.UserRole, workbook: ExcelJS.Workbook): Promise<void> {
    await entity.load().then(loadedEntity =>
        checkIfModuleRoleIsUsedForEntityRole(loadedEntity, loadedEntity.accessRules, moduleRole, userRole, workbook));
}

async function processAllPages(modRole: security.IModuleRole, userRole: security.UserRole, workbook: ExcelJS.Workbook): Promise<void> {
    await Promise.all(modRole.model.allPages().map(async (page) => processPage(modRole, userRole, page, workbook)));
}

async function processPage(modRole: security.IModuleRole, userRole: security.UserRole, page: pages.IPage, workbook: ExcelJS.Workbook): Promise<void> {
    await page.load().then(loadedPage => addPage(modRole, userRole, loadedPage, workbook));
}

function addPage(modRole: security.IModuleRole, userRole: security.UserRole, loadedPage: pages.Page, workbook: ExcelJS.Workbook) {
    const allowed = loadedPage.allowedRoles.filter(allowedRole => allowedRole.name == modRole.name).length > 0;
    getWorksheet(workbook, 'Pages').addRow([
        truncateString(userRole.name),
        truncateString(modRole.containerAsModuleSecurity.containerAsModule.name),
        truncateString(modRole.name),
        truncateString(loadedPage.name),
        allowed ? 'True' : 'False'
    ]);
}


///section to process microflows
async function processAllMicroflows(modRole: security.IModuleRole, userRole: security.UserRole): Promise<void> {
    await Promise.all(modRole.model.allMicroflows().map(async (microflow) => processMicroflow(modRole, userRole, microflow)));
}

async function processMicroflow(modRole: security.IModuleRole, userRole: security.UserRole, microflow: microflows.IMicroflow, workbook: ExcelJS.Workbook): Promise<void> {
    await microflow.load().then(microflowLoaded => addMicroflow(modRole, userRole, microflowLoaded, workbook));
}
function addMicroflow(modRole: security.IModuleRole, userRole: security.UserRole, microflowLoaded: microflows.Microflow, workbook: ExcelJS.Workbook) {
    const allowed = microflowLoaded.allowedModuleRoles.filter(allowedRole => allowedRole.name == modRole.name).length > 0;
    getWorksheet(workbook, 'Microflows').addRow([
        truncateString(userRole.name),
        truncateString(modRole.containerAsModuleSecurity.containerAsModule.name),
        truncateString(modRole.name),
        truncateString(microflowLoaded.name),
        allowed ? 'True' : 'False'
    ]);
}

///section to process nanoflows
async function processAllNanoflows(modRole: security.IModuleRole, userRole: security.UserRole): Promise<void> {
    await Promise.all(modRole.model.allNanoflows().map(async (nanoflow) => processNanoflow(modRole, userRole, nanoflow)));
}

async function processNanoflow(modRole: security.IModuleRole, userRole: security.UserRole, nanoflow: microflows.INanoflow, workbook: ExcelJS.Workbook): Promise<void> {
    await nanoflow.load().then(nanoflowLoaded => addNanoflow(modRole, userRole, nanoflowLoaded, workbook));
}
function addNanoflow(modRole: security.IModuleRole, userRole: security.UserRole, nanoflowLoaded: microflows.Nanoflow, workbook: ExcelJS.Workbook) {
    const allowed = nanoflowLoaded.allowedModuleRoles.filter(allowedRole => allowedRole.name == modRole.name).length > 0;
    getWorksheet(workbook, 'Nanoflows').addRow([
        truncateString(userRole.name),
        truncateString(modRole.containerAsModuleSecurity.containerAsModule.name),
        truncateString(modRole.name),
        truncateString(nanoflowLoaded.name),
        allowed ? 'True' : 'False'
    ]);
}


async function checkIfModuleRoleIsUsedForEntityRole(entity: domainmodels.Entity, accessRules: domainmodels.AccessRule[], modRole: security.IModuleRole, userRole: security.UserRole, workbook: ExcelJS.Workbook): Promise<void> {
    await Promise.all(accessRules.map(async (rule) => {
        if (rule.moduleRoles.filter(entityModRule => entityModRule.name === modRole.name).length > 0) {
            let memberRules = rule.memberAccesses.reduce((acc, memRule) => {
                if (memRule && memRule.accessRights && memRule.attribute) {
                    return acc + `${memRule.attribute.name}: ${memRule.accessRights.name}\n`;
                }
                return acc;
            }, '');

            let createDelete = 'None';
            if (rule.allowCreate && rule.allowDelete) createDelete = 'Create/Delete';
            else if (rule.allowCreate) createDelete = 'Create';
            else if (rule.allowDelete) createDelete = 'Delete';

            getWorksheet(workbook, 'Entities').addRow([
                truncateString(userRole.name),
                truncateString(entity.containerAsDomainModel.containerAsModule.name),
                truncateString(modRole.name),
                truncateString(entity.name),
                truncateString(rule.xPathConstraint),
                createDelete,
                truncateString(memberRules)
            ]);
        }
    }));
}

function addIfModuleRoleInUserRole(modRole: security.IModuleRole, userRole: security.UserRole): boolean {
    // console.debug(`Processing module role: ${modRole.name}`);
    if (userRole.moduleRoles.filter(modRoleFilter => {
        if (modRoleFilter != null) {
            return modRoleFilter.name === modRole.name;
        } else {
            return false;
        }
    }).length > 0) {
        return true;
    } else {
        return false;
    }

}

/**
* This function loads the project security.
*/
async function loadProjectSecurity(workingCopy: OnlineWorkingCopy): Promise<security.ProjectSecurity> {

    var model: IModel = await workingCopy.openModel();
    var security = model.allProjectSecurities()[0];
    return await security.load();
}

function getAllUserRoles(projectSecurity: security.ProjectSecurity): security.UserRole[] {
    console.log("All user roles retrieved");
    return projectSecurity.userRoles;
}
