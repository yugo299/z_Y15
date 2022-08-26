const vCat = '20'; //ゲーム ■■■■■■
const nNum = 50, rNum = 100, iNum = 21; //次の指標を見るための加算行数、ランク数、指標の数

let tmpM = new Date();
tmpM.setDate(tmpM.getDate()+5);
const nMonth = Utilities.formatDate(tmpM, 'Etc/GMT-4', 'yyMM');
const tMonth = Utilities.formatDate(new Date(), 'Etc/GMT-4', 'yyMM');
const today = Utilities.formatDate(new Date(), 'Etc/GMT-4', 'dd'); //g-P_CAT_YM のシート名

const tmpP_ID = '1KZqw7EQMfwCHq2FdA_ycxxrc8PvBGEupYeHTIlw-ORU'; //g-P_CAT_YM ■■■■■■
const tmpV_ID = '1v3WA7PmX9qWFaSrDpu2yTUbMcL5mgAF28NuLVcyPOZk'; //gYM_V CAT ■■■■■■
const tmpC_ID = '15DdAxfrXA4MmWd7RvHkKIToyD1VQee5KPDX3gBO4lfc'; //gYM_C CAT ■■■■■■

const tmpP = DriveApp.getFileById(tmpP_ID);
const tmpV = DriveApp.getFileById(tmpV_ID);
const tmpC = DriveApp.getFileById(tmpC_ID);

const pFolder_ID = '1JzikFpMyzb_NO9iIbZV1cmsEGzDr8eMS'; //y20 ■■■■■■
const pFolder = DriveApp.getFolderById(pFolder_ID);

const cFolder_ID = DriveApp.getFoldersByName('y' + vCat + '_' + tMonth).next().getId();
const cFolder = DriveApp.getFolderById(cFolder_ID);

const pFile_ID = cFolder.getFilesByName('gP' + vCat + '_' + tMonth).next().getId();
const iFile_ID = cFolder.getFilesByName('g'+ tMonth + today + '_I' + vCat).next().getId();
const cFile_ID = cFolder.getFilesByName('g'+ tMonth + today + '_C' + vCat).next().getId();

const pFile = SpreadsheetApp.openById(pFile_ID);
const iFile = SpreadsheetApp.openById(iFile_ID);
const cFile = SpreadsheetApp.openById(cFile_ID);

const pSheet = pFile.getSheetByName(today);

const vISheet = iFile.getSheetByName('vI');
const vRtSheet = iFile.getSheetByName('vRt');
const vRnSheet = iFile.getSheetByName('vRn');
const vVSheet = iFile.getSheetByName('vV');
const vLSheet = iFile.getSheetByName('vL');
const vCSheet = iFile.getSheetByName('vC');
const cISheet = iFile.getSheetByName('cI');
const cRtSheet = iFile.getSheetByName('cRt');
const cRnSheet = iFile.getSheetByName('cRn');
const cVSheet = iFile.getSheetByName('cV');
const cLSheet = iFile.getSheetByName('cL');
const cCSheet = iFile.getSheetByName('cC');
const cSSheet = iFile.getSheetByName('cS');
const cNSheet = iFile.getSheetByName('cN');
const cTSheet = iFile.getSheetByName('cT');

let tmpD = new Date(Utilities.formatDate(new Date(), 'Etc/GMT-4', 'yyyy/MM/dd-HH:mm'));
const tRow = 3 + Math.floor(tmpD.getMinutes()/30) + tmpD.getHours()*2;