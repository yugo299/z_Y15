const vCat = '2'; //自動車と乗り物 ■■■■■■
const nNum = 50, rNum = 100, iNum = 21; //次の指標を見るための加算行数、ランク数、指標の数

let tmpM = new Date();
tmpM.setDate(tmpM.getDate()+5);
const nMonth = Utilities.formatDate(tmpM, 'Etc/GMT-4', 'yyMM');
const tMonth = Utilities.formatDate(new Date(), 'Etc/GMT-4', 'yyMM');
const today = Utilities.formatDate(new Date(), 'Etc/GMT-4', 'dd'); //gP_CAT_YM のシート名
const rDay = JSON.stringify(Utilities.formatDate(new Date(), 'JST', 'yy-MM-dd-')); //ランキング用日時

const tmpP_ID = '1KZqw7EQMfwCHq2FdA_ycxxrc8PvBGEupYeHTIlw-ORU'; //g-PCAT_YM ■■■■■■
const tmpR_ID = '1t1lO_eUlxNnlAx8acyL9r2PmgwQIUDXD-0SSGZ2vaHQ'; //gYM_R CAT ■■■■■■

const tmpP = DriveApp.getFileById(tmpP_ID);
const tmpR = DriveApp.getFileById(tmpR_ID);

const pFolder_ID = '1K1E0GUfEu9d7jAehkSw8lUP7y65ASQ9B'; //y02 ■■■■■■
const pFolder = DriveApp.getFolderById(pFolder_ID);

const cFolder_ID = DriveApp.getFoldersByName('y' + vCat + '_' + tMonth).next().getId();
const cFolder = DriveApp.getFolderById(cFolder_ID);

const pFile_ID = cFolder.getFilesByName('gP' + vCat + '_' + tMonth).next().getId();
const rFile_ID = cFolder.getFilesByName('g'+ tMonth + today + '_R' + vCat).next().getId();

const pFile = SpreadsheetApp.openById(pFile_ID);
const rFile = SpreadsheetApp.openById(rFile_ID);

const pSheet = pFile.getSheetByName(today);

const vISheet = rFile.getSheetByName('vI');
const vRtSheet = rFile.getSheetByName('vRt');
const vRnSheet = rFile.getSheetByName('vRn');
const vPSheet = rFile.getSheetByName('vP');
const vVSheet = rFile.getSheetByName('vV');
const vLSheet = rFile.getSheetByName('vL');
const vCSheet = rFile.getSheetByName('vC');
const cISheet = rFile.getSheetByName('cI');
const cRtSheet = rFile.getSheetByName('cRt');
const cRnSheet = rFile.getSheetByName('cRn');
const cPSheet = rFile.getSheetByName('cP');
const cVSheet = rFile.getSheetByName('cV');
const cLSheet = rFile.getSheetByName('cL');
const cCSheet = rFile.getSheetByName('cC');
const cSSheet = rFile.getSheetByName('cS');
const cNSheet = rFile.getSheetByName('cN');
const cTSheet = rFile.getSheetByName('cT');

let tmpD = new Date(Utilities.formatDate(new Date(), 'Etc/GMT-4', 'yyyy/MM/dd-HH:mm'));
const tRow = 3 + Math.floor(tmpD.getMinutes()/30) + tmpD.getHours()*2;

const ratio = [
  746,773,800,827,854,881,908,935,962,989,
  1016,1043,1070,1097,1124,1151,1178,1205,1232,1259,
  1286,1313,1340,1367,1394,1421,1448,1475,1502,1529,
  1556,1583,1610,1637,1664,1691,1718,1745,1772,1799,
  1826,1853,1880,1907,1934,1961,1988,2015,2042,2069,
  2097,2124,2151,2178,2205,2232,2259,2286,2313,2340,
  2367,2394,2421,2448,2475,2502,2529,2556,2583,2610,
  2637,2664,2691,2718,2745,2772,2799,2826,2853,2880,
  2908,2935,2962,2989,3016,3043,3070,3097,3124,3151,
  3179,3206,3233,3260,3287,3314,3341,3369,3396,3424
]