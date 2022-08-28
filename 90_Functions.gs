function writeIndicator (cat, dI, dRn, dRt, dP, dV, dL, dC, dS, dN, dT) {

  let row = dI.length;
  switch (cat) {
    case 'v':
      vISheet.getRange(1, 1, row, 10).setValues(dI);
      vRnSheet.getRange(1, 1, row, 148).setValues(dRn);
      vRtSheet.getRange(1, 1, row, 196).setValues(dRt);
      vPSheet.getRange(1, 1, row, 196).setValues(dP);
      vVSheet.getRange(1, 1, row, 148).setValues(dV);
      vLSheet.getRange(1, 1, row, 148).setValues(dL);
      vCSheet.getRange(1, 1, row, 148).setValues(dC);
      return;
    case 'c':
      cISheet.getRange(1, 1, row, 8).setValues(dI);
      cRnSheet.getRange(1, 1, row, 148).setValues(dRn);
      cRtSheet.getRange(1, 1, row, 196).setValues(dRt);
      cPSheet.getRange(1, 1, row, 196).setValues(dP);
      cVSheet.getRange(1, 1, row, 148).setValues(dV);
      cLSheet.getRange(1, 1, row, 148).setValues(dL);
      cCSheet.getRange(1, 1, row, 148).setValues(dC);
      cSSheet.getRange(1, 1, row, 99).setValues(dS);
      cNSheet.getRange(1, 1, row, 99).setValues(dN);
      cTSheet.getRange(1, 1, row, 99).setValues(dT);
      return;
    default: console.log('■■■■■■■■■■  エラー : writeIndicator  ■■■■■■■■■■')
  }
}