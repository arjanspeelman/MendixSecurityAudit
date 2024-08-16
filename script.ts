import { MendixPlatformClient, OnlineWorkingCopy } from "mendixplatformsdk";
import { ModelSdkClient, IModel, projects, domainmodels, microflows, pages, navigation, texts, security, IStructure, menus, IList, nanoflows } from "mendixmodelsdk";
import * as fs from 'fs';
import * as stream from 'stream';
import * as officegen from 'officegen';
import * as ExcelJS from 'exceljs';

// const appId = "{{AppID}}";
const appId = "c99be9a4-8ccf-4c29-aabb-7ea0c7242ebc";
const branchName = null // null for mainline
const wc = null;
const client = new MendixPlatformClient();
const workbook = new ExcelJS.Workbook();
let pObj;

const CHUNK_SIZE = 5000; // Verhoogd aantal items om per keer te verwerken

// Functie om lange strings af te kappen
function truncateString(str: string, maxLength: number = 32000): string {
    if (str && str.length > maxLength) {
        return str.substring(0, maxLength) + '...';
    }
    return str;
}

const sheetEntities = workbook.addWorksheet('Entities');
sheetEntities.addRow(['User Role', 'Module', 'Module Role', 'Entity', 'Xpath', 'Create/Delete', 'Member Rules']);

const sheetPages = workbook.addWorksheet('Pages');
sheetPages.addRow(['User Role', 'Module', 'Module Role', 'Page Name', 'Allowed']);

const sheetMicroflows = workbook.addWorksheet('Microflows');
sheetMicroflows.addRow(['User Role', 'Module', 'Module Role', 'Microflows', 'Allowed']);

const sheetNanoflows = workbook.addWorksheet('Nanoflows');
sheetNanoflows.addRow(['User Role', 'Module', 'Module Role', 'Nanoflows', 'Allowed']);

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

    await createUserSecurityDocument(userRoles);

    await workbook.xlsx.writeFile('MendixSecurityDocument.xlsx');
    console.log('Finished creating Document');
    
    // Forceer garbage collection
    if (global.gc) {
        global.gc();
    }
}

/**
* This function picks the first navigation document in the project.
*/
async function createUserSecurityDocument(userRoles: security.UserRole[]) {
    console.log("Creating User Access Matrix");
    for (let i = 0; i < userRoles.length; i += CHUNK_SIZE) {
        const chunk = userRoles.slice(i, i + CHUNK_SIZE);
        await Promise.all(chunk.map(async (userRole) => processAllModules(userRole)));
    }
}

async function processAllModules(userRole: security.UserRole): Promise<void> {
    const modules = userRole.model.allModules();
    for (let i = 0; i < modules.length; i += CHUNK_SIZE) {
        const chunk = modules.slice(i, i + CHUNK_SIZE);
        await Promise.all(chunk.map(async (module) => processModule(module, userRole)));
    }
}

async function processModule(module: projects.IModule, userRole: security.UserRole): Promise<void> {
    // console.debug(`Processing module: ${module.name}`);
    var securities = await getAllModuleSecurities(module);
    await Promise.all(securities.map(async (security) => loadAllModuleSecurities(securities, userRole)));

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

async function loadAllModuleSecurities(moduleSecurities: security.IModuleSecurity[], userRole: security.UserRole): Promise<void> {
    await Promise.all(moduleSecurities.map(async (mSecurity) => processLoadedModSec(mSecurity, userRole)));
}

async function processLoadedModSec(modSec: security.IModuleSecurity, userRole: security.UserRole): Promise<void> {
    await Promise.all(modSec.moduleRoles.map(async (modRole) => processModRole(modRole, userRole)));
}



async function loadModSec(modSec: security.IModuleSecurity): Promise<security.ModuleSecurity> {
    // console.debug(`Processing loadModSec`);
    return modSec.load();
}



async function processModRole(modRole: security.IModuleRole, userRole: security.UserRole): Promise<void> {
    if (addIfModuleRoleInUserRole(modRole, userRole)) {
        await Promise.all(modRole.containerAsModuleSecurity.containerAsModule.domainModel.entities.map(async (entity) =>
            processAllEntitySecurityRules(entity, modRole, userRole)
                .then(() => processAllPages(modRole, userRole))
                .then(() => processAllMicroflows(modRole, userRole))
                .then(() => processAllNanoflows(modRole, userRole))));
    }

}
async function processAllEntitySecurityRules(entity: domainmodels.IEntity, moduleRole: security.IModuleRole, userRole: security.UserRole): Promise<void> {
    await entity.load().then(loadedEntity =>
        checkIfModuleRoleIsUsedForEntityRole(loadedEntity, loadedEntity.accessRules, moduleRole, userRole));
}

async function processAllPages(modRole: security.IModuleRole, userRole: security.UserRole): Promise<void> {
    await Promise.all(modRole.model.allPages().map(async (page) => processPage(modRole, userRole, page)));
}

async function processPage(modRole: security.IModuleRole, userRole: security.UserRole, page: pages.IPage): Promise<void> {
    await page.load().then(loadedPage => addPage(modRole, userRole, loadedPage));
}

function addPage(modRole: security.IModuleRole, userRole: security.UserRole, loadedPage: pages.Page) {
    const allowed = loadedPage.allowedRoles.filter(allowedRole => allowedRole.name == modRole.name).length > 0;
    sheetPages.addRow([
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

async function processMicroflow(modRole: security.IModuleRole, userRole: security.UserRole, microflow: microflows.IMicroflow): Promise<void> {
    await microflow.load().then(microflowLoaded => addMicroflow(modRole, userRole, microflowLoaded));
}
function addMicroflow(modRole: security.IModuleRole, userRole: security.UserRole, microflowLoaded: microflows.Microflow) {
    const allowed = microflowLoaded.allowedModuleRoles.filter(allowedRole => allowedRole.name == modRole.name).length > 0;
    sheetMicroflows.addRow([
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

async function processNanoflow(modRole: security.IModuleRole, userRole: security.UserRole, nanoflow: microflows.INanoflow): Promise<void> {
    await nanoflow.load().then(nanoflowLoaded => addNanoflow(modRole, userRole, nanoflowLoaded));
}
function addNanoflow(modRole: security.IModuleRole, userRole: security.UserRole, nanoflowLoaded: microflows.Nanoflow) {
    const allowed = nanoflowLoaded.allowedModuleRoles.filter(allowedRole => allowedRole.name == modRole.name).length > 0;
    sheetNanoflows.addRow([
        truncateString(userRole.name),
        truncateString(modRole.containerAsModuleSecurity.containerAsModule.name),
        truncateString(modRole.name),
        truncateString(nanoflowLoaded.name),
        allowed ? 'True' : 'False'
    ]);
}


async function checkIfModuleRoleIsUsedForEntityRole(entity: domainmodels.Entity, accessRules: domainmodels.AccessRule[], modRole: security.IModuleRole, userRole: security.UserRole): Promise<void> {
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

            sheetEntities.addRow([
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
