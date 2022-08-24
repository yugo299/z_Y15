const vCat = '20'; //ゲーム ■■■■■■

let tmpM = new Date();
tmpM.setDate(tmpM.getDate()+5);
const nMonth = Utilities.formatDate(tmpM, 'Asia/Baku', 'yyMM');
const tMonth = Utilities.formatDate(new Date(), 'Asia/Baku', 'yyMM');
const today = Utilities.formatDate(new Date(), 'Asia/Baku', 'dd'); //g-P_CAT_YM のシート名

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

const pFile_ID = cFolder.getFilesByName('g-P' + vCat + '_' + tMonth).next().getId();
const vFile_ID = cFolder.getFilesByName('g'+ tMonth + today + '_V' + vCat).next().getId();
const cFile_ID = cFolder.getFilesByName('g'+ tMonth + today + '_C' + vCat).next().getId();

const pFile = SpreadsheetApp.openById(pFile_ID);
const vFile = SpreadsheetApp.openById(vFile_ID);
const cFile = SpreadsheetApp.openById(cFile_ID);

let tmpD = new Date(Utilities.formatDate(new Date(), 'Asia/Baku', 'yy/MM/dd-HH:mm'));
const firstRow = 2 + Math.floor(tmpD.getMinutes()/15) + tmpD.getHours()*4;