const vCat = '2'; //自動車と乗り物 ■■■■■■
const nNum = 50, rNum = 100, iNum = 21; //次の指標を見るための加算行数、ランク数、指標の数

let tmpD = new Date();
tmpD.setDate(tmpD.getDate()+5);
const nMonth = Utilities.formatDate(tmpD, 'Etc/GMT-4', 'yyMM');
const tMonth = Utilities.formatDate(new Date(), 'Etc/GMT-4', 'yyMM');
const today = Utilities.formatDate(new Date(), 'Etc/GMT-4', 'dd'); //gP_CAT_YM のシート名
const rDay = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd-'); //ランキング用日時

tmpD = new Date();
tmpD.setDate(tmpD.getDate()+1);
const tomorrow = Utilities.formatDate(tmpD, 'Etc/GMT-4','dd');

tmpD = new Date(Utilities.formatDate(new Date(), 'Etc/GMT-4', 'yyyy/MM/dd-HH:mm'));
const tRow = 3 + Math.floor(tmpD.getMinutes()/30) + tmpD.getHours()*2;

const tmpP_ID = '1KZqw7EQMfwCHq2FdA_ycxxrc8PvBGEupYeHTIlw-ORU'; //g-PCAT_YM ■■■■■■
const tmpR_ID = '1t1lO_eUlxNnlAx8acyL9r2PmgwQIUDXD-0SSGZ2vaHQ'; //gYM_R CAT ■■■■■■

const tmpP = DriveApp.getFileById(tmpP_ID);
const tmpR = DriveApp.getFileById(tmpR_ID);

const gFolder = DriveApp.getFolderById('14u0G2CGKp3TOYkOYGRcacZ25wY5vUDls');

const pFolder_ID = gFolder.getFoldersByName('y' + vCat).next().getId();
const pFolder = DriveApp.getFolderById(pFolder_ID);

const cFolder_ID = pFolder.getFoldersByName('y' + vCat + '_' + tMonth).next().getId();
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

const ratio = 	[
  3424,3396,3369,3341,3314,3287,3260,3233,3206,3179,
  3151,3124,3097,3070,3043,3016,2989,2962,2935,2908,
  2880,2853,2826,2799,2772,2745,2718,2691,2664,2637,
  2610,2583,2556,2529,2502,2475,2448,2421,2394,2367,
  2340,2313,2286,2259,2232,2205,2178,2151,2124,2097,
  2069,2042,2015,1988,1961,1934,1907,1880,1853,1826,
  1799,1772,1745,1718,1691,1664,1637,1610,1583,1556,
  1529,1502,1475,1448,1421,1394,1367,1340,1313,1286,
  1259,1232,1205,1178,1151,1124,1097,1070,1043,1016,
  989,962,935,908,881,854,827,800,773,746
]

let newC = new Set();