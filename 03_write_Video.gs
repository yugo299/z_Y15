function functionV() {
  if (pSheet.getRange(tRow, 1).getValue()!='popular') { return }

  let row = vISheet.getLastRow();
  let vI = vISheet.getRange(1, 1, row, 10).getValues();
  let vRt = vRtSheet.getRange(1, 1, row, 196).getValues();
  let vV = vVSheet.getRange(1, 1, row, 148).getValues();
  let vL = vLSheet.getRange(1, 1, row, 148).getValues();
  let vC = vCSheet.getRange(1, 1, row, 148).getValues();

  let vRn = vRnSheet.getRange(2, 1, row-1, 148).getValues();
  let arr = vRnSheet.getRange(1, 1, 1, 148).getDisplayValues()[0];
  vRn.unshift(arr);

  let vP = vPSheet.getRange(2, 1, row-1, 196).getValues();
  arr = vPSheet.getRange(1, 1, 1, 196).getDisplayValues()[0];
  vP.unshift(arr);

  for (let i=1; i<vI.length; i++) {
    vRt[i][2] = '', vRt[i][3] = '', vRt[i][52+tRow-3] = vRt[i][52+tRow-3-1];
    vV[i][2] = '', vV[i][3] = '';
    vL[i][2] = '', vL[i][3] = '';
    vC[i][2] = '', vC[i][3] = '';
  }

  let src = Array(iNum+1);

  row = tRow;
  for (let i=0; i<iNum; i++) {
    src[i] = pSheet.getRange(row, 3, 1, rNum).getValues()[0];
    row += nNum;
  }
  src[iNum] = pSheet.getRange(1, 3, 1, rNum).getValues()[0];
  src = src[0].map((_, c) => src.map(r => r[c]));

  for (let i=0; i<rNum; i++) {
    row = vI.findIndex(x => x[0] === src[i][0]);
    vI = vInfo(row, src[i], vI);
    vRn = vRank(row, src[i], vRn, iNum);
    vRt = vRatio(row, src[i], vRt);
    vP = vPeriod(row, src[i], vP);
    vV = vOther(row, src[i], vV, 4);
    vL = vOther(row, src[i], vL, 5);
    vC = vOther(row, src[i], vC, 6);
  }

  writeIndicator('v', vI, vRn, vRt, vP, vV, vL, vC);

  pSheet.getRange(tRow, 1).setValue('video')
}

//  vID, vTitle, vDate, dur, cntV, cntL, cntC, cID, cTitle, cDate,
//  sub, cntN, totV, vDesc, vURL, vTmb, vTags, cDesc, cURL, cTmb,
//  cCustom, rank


function vInfo(row, src, data) {

  let d = Array(10);
  d[0] = src[0];
  d[1] = src[1];
  d[2] = src[7];
  d[3] = src[8];
  d[4] = src[2];
  d[5] = src[3];
  d[6] = src[13];
  d[7] = src[14];
  d[8] = src[15];
  d[9] = src[16];

  if (row != -1) { data[row] = d }
  else { data.push(d) }

  return data
}

function vRank(row, src, data, clm) {

  if (row != -1) {
    data[row][0] = src[0];
    data[row][1] = src[1];
    data[row][52+tRow-3] = src[clm];
    //前日との差分
    if (data[row][52+tRow-3-48]) {
      data[row][52+tRow-3+48] = data[row][52+tRow-3] - data[row][52+tRow-3-48]
    }
    //最高順位、日時
    if (src[clm] < data[row][2]) {
      data[row][2] = src[clm];
      data[row][3] = rDay + data[0][52+tRow-3];
    }
  }
  else {
    let d = Array(148);
    d[0] = src[0];
    d[1] = src[1];
    d[2] = src[clm];
    d[3] = rDay + data[0][52+tRow-3];
    d[52+tRow-3] = src[clm];
    data.push(d);
  }
  return data
}

function vRatio(row, src, data) {

  if (row != -1) {
    data[row][1] = src[1];
    data[row][148+tRow-3] = ratio[src[iNum]-1];
    data[row][52+tRow-3] = data[row][52+tRow-3-1] + ratio[src[iNum]-1];
    //直近1h
    if (data[row][52+tRow-3-2]) {
      data[row][2] = data[row][52+tRow-3] - data[row][52+tRow-3-2]
    }
    //直近24h
    if (data[row][52+tRow-3-48]) {
      data[row][3] = data[row][52+tRow-3] - data[row][52+tRow-3-48]
    }
  }
  else {
    let d = Array(196);
    d[0] = src[0];
    d[1] = src[1];
    d[148+tRow-3] = ratio[src[iNum]-1];
    d[52+tRow-3] = ratio[src[iNum]-1];
    data.push(d);
  }
  return data
}

function vPeriod(row, src, data) {

  if (row != -1) {
    data[row][1] = src[1];
    //ランクイン開始日時
    if (data[row][148+tRow-3-1]) { data[row][148+tRow-3] = data[row][148+tRow-3-1] }
    else { data[row][148+tRow-3] = rDay + data[0][148+tRow-3] }
    //ランクイン期間
    if (data[row][52+tRow-3-1]) { data[row][52+tRow-3] = data[row][52+tRow-3-1] + 0.5 }
    else { data[row][52+tRow-3] = 0.5 }
    //最長期間の更新
    if (data[row][52+tRow-3] > data[row][2]) {
      data[row][2] = data[row][52+tRow-3];
      data[row][3] = data[row][52+tRow-3] + '～' + rDay + data[0][52+tRow-3];
    }
  }
  else {
    let d = Array(196);
    d[0] = src[0];
    d[1] = src[1];
    d[2] = 0.5;
    d[3] = rDay + data[0][52+tRow-3] + '～' + rDay + data[0][52+tRow-3];
    d[52+tRow-3] = 0.5;
    d[52+tRow-3+96] = rDay + data[0][52+tRow-3+96];
    data.push(d);
  }
  return data
}

function vOther(row, src, data, clm) {

  if (row != -1) {
    data[row][1] = src[1];
    data[row][52+tRow-3] = src[clm];
    //前回との差分
    if (data[row][52+tRow-3-1]) {
      data[row][52+tRow-3+48] = data[row][52+tRow-3] - data[row][52+tRow-3-1]
    }
    //直近1h
    if (data[row][52+tRow-3-2]) {
      data[row][2] = data[row][52+tRow-3] - data[row][52+tRow-3-2]
    }
    //直近24h
    if (data[row][52+tRow-3-48]) {
      data[row][3] = data[row][52+tRow-3] - data[row][52+tRow-3-48]
    }
  }
  else {
    let d = Array(148);
    d[0] = src[0];
    d[1] = src[1];
    d[52+tRow-3] = src[clm];
    data.push(d);
  }
  return data
}