function createMonthly() {

  let name = "y" + vCat + '_' + nMonth;
  pFolder.createFolder(name);

  const nFolderID = DriveApp.getFoldersByName(name).next().getId();
  const nFolder = DriveApp.getFolderById(nFolderID);

  name = 'gP' + vCat + '_' + nMonth;
  tmpP.makeCopy(name, nFolder);

  name = 'g' + nMonth + '01_R' + vCat;
  tmpR.makeCopy(name, nFolder);
}

function createDaily() {

  let tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate()+1);
  tomorrow = Utilities.formatDate(tomorrow, 'Asia/Baku','dd');
  if (tomorrow === '01') { return }

  const copySheet = pFile.getSheetByName('dd');
  const newSheet = copySheet.copyTo(pFile);
  newSheet.setName(tomorrow);

  let name = 'g'+ tMonth + tomorrow + '_R' + vCat;
  tmpR.makeCopy(name, cFolder);
}