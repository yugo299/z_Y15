function functionGG() {
  if (checkPopular()) { return }
  functionP();
  functionV();
  functionC();
  createDaily();
  createMonthly();
  transferDaily();
}

//■■■■■■■■■ 急上昇 ■■■■■■■■■

function functionP() {

  const srcData1 = getPopular('');
  const srcData2 = getPopular(srcData1[0]);

//  nextPageToken, vID, vTitle, vDate, dur, cntV, cntL, cntC, cID, cTitle,
//  cDate, sub, cntN, totV, vDesc, vURL, vTmb, vTags, cDesc, cURL,
//  cTmb, cCustom

  let srcData = [...Array(srcData1.length-1)].map(() => []);

  for (let i=0; i<srcData.length; i++) {
    srcData[i] = srcData1[i+1].concat(srcData2[i+1])
  }

  writePopular(srcData);

  pSheet.getRange(tRow, 1).setValue('popular')
}

function checkPopular() {
  const time = pSheet.getRange(tRow, 2).getValue();
  const data = pSheet.getRange(tRow, 3).getValue();
  if (data) {
    console.log('実施済み : ' + time);
    return true
  }
  return false
}

function getPopular(nextPageToken) {

  let [
    token, vID, vTitle, vDate, dur, cntV, cntL, cntC, cID, cTitle,
    cDate, cntS, cntN, totV, vDesc, vURL, vTmb, vTags, cDesc, cURL,
    cTmb, cCustom
  ] = [
    '',[],[],[],[],[],[],[],[],[],
    [],[],[],[],[],[],[],[],[],[],
    [],[]
  ];
  let c_tmp = ''

  const vfields = 'items(id,snippet(title,description,publishedAt,thumbnails(medium(url)),tags,channelId),contentDetails(duration),statistics(viewCount,likeCount,commentCount)),nextPageToken';
  let optJson = {chart: 'mostPopular', regionCode: 'jp', videoCategoryId: vCat, maxResults: 50, fields: vfields, pageToken: nextPageToken};

  vJson = YouTube.Videos.list('snippet,contentDetails,statistics',optJson);
  token = vJson.nextPageToken;

  vJson.items.forEach((vJ) => {

    const cfields = 'items(id,snippet(title,description,publishedAt,thumbnails(medium(url)),customUrl),statistics(viewCount,subscriberCount,videoCount)),nextPageToken';
    optJson = {id: vJ.snippet.channelId, fields: cfields};
    cJ = YouTube.Channels.list('snippet,statistics',optJson).items[0];

    vID.push(vJ.id);
    vTitle.push(vJ.snippet.title);
    vDate.push(Utilities.formatDate(new Date(vJ.snippet.publishedAt), 'JST', 'yyyy-MM-dd HH:mm:ss'));
    dur.push(convertTime(vJ.contentDetails.duration));
    cntV.push(vJ.statistics.viewCount);
    cntL.push(vJ.statistics.likeCount);
    cntC.push(vJ.statistics.commentCount);
    cID.push(vJ.snippet.channelId);
    cTitle.push(cJ.snippet.title);
    cDate.push(Utilities.formatDate(new Date(cJ.snippet.publishedAt), 'JST"', 'yyyy-MM-dd HH:mm:ss'));
    cntS.push(cJ.statistics.subscriberCount);
    cntN.push(cJ.statistics.videoCount);
    totV.push(cJ.statistics.viewCount);
    vDesc.push(vJ.snippet.description);
    vURL.push('https://youtube.com/watch?v='+vJ.id);
    vTmb.push(vJ.snippet.thumbnails.medium.url);
    vTags.push(joinArr(vJ.snippet.tags));
    cDesc.push(cJ.snippet.description);
    cTmb.push(cJ.snippet.thumbnails.medium.url);

    c_tmp = removeAt(cJ.snippet.customUrl)
    cCustom.push(c_tmp);

    if (c_tmp) { cURL.push('https://youtube.com/c/'+c_tmp) }
    else { cURL.push('https://youtube.com/channel/'+vJ.snippet.channelId) }
  })

  const srcData = [
    token, vID, vTitle, vDate, dur, cntV, cntL, cntC, cID, cTitle,
    cDate, cntS, cntN, totV, vDesc, vURL, vTmb, vTags, cDesc, cURL,
    cTmb, cCustom
  ];

  return srcData
}

function writePopular(srcData) {

  let row = tRow;
  for (let i=0; i<srcData.length; i++) {
    pSheet.getRange(row, 3, 1, rNum).setValues([srcData[i]]);
    row += nNum;
  }
}

function joinArr(arr) {

  if (typeof(arr)==='object') { return arr.join(',') }
}

function removeAt(str) {

  if (!str) { return }
  if (str.slice(0,6) === '@user-') { return }
  if (str.slice(0,1) != '@') { return }
  return str.slice(1)
}

function convertTime(duration) {

  if (!duration || duration==='P0D') { return }
  var reg = new RegExp('^PT([0-9]*H)?([0-9]*M)?([0-9]*S)?');
  var regResult = duration.match(reg);

  var hour = regResult[1];
  var minutes = regResult[2];
  var sec = regResult[3];

  if(hour == undefined) {hour = '00';}
  else {
    hour = hour.split('H')[0];
    if(hour.length == 1){hour = '0' + hour;}
  }

  if(minutes == undefined) {minutes = '00';}
  else {
    minutes = minutes.split('M')[0];
    if(minutes.length == 1){minutes = '0' + minutes;}
  }

  if(sec == undefined) {sec = '00';}
  else {
    sec = sec.split('S')[0];
    if(sec.length == 1){sec = '0' + sec;}
  }

  return hour + ":" + minutes + ":" + sec
}

//■■■■■■■■■ 動画詳細 ■■■■■■■■■

function functionV() {

  let row = vISheet.getLastRow();
  let vI = vISheet.getRange(1, 1, row, 10).getValues();
  let vRt = vRtSheet.getRange(1, 1, row, 196).getValues();
  let vV = vVSheet.getRange(1, 1, row, 148).getValues();
  let vL = vLSheet.getRange(1, 1, row, 148).getValues();
  let vC = vCSheet.getRange(1, 1, row, 148).getValues();

  let vRn = [];
  let arr = vRnSheet.getRange(1, 1, 1, 148).getDisplayValues();
  if (row > 1) {
    vRn = vRnSheet.getRange(2, 1, row-1, 148).getValues();
    vRn.unshift(arr[0]);
  }
  else { vRn = arr }

  let vP = [];
  arr = vPSheet.getRange(1, 1, 1, 196).getDisplayValues();
  if (row > 1) {
    vP = vPSheet.getRange(2, 1, row-1, 196).getValues();
    vP.unshift(arr[0]);
  }
  else { vP = arr }

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
      data[row][3] = data[row][148+tRow-3] + '～' + rDay + data[0][52+tRow-3];
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
    data[row][0] = src[0];
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

//■■■■■■■■■ チャンネル詳細 ■■■■■■■■■

function functionC() {

  let row = cISheet.getLastRow();
  let cI = cISheet.getRange(1, 1, row, 8).getValues();
  let cRt = cRtSheet.getRange(1, 1, row, 196).getValues();
  let cV = cVSheet.getRange(1, 1, row, 148).getValues();
  let cL = cLSheet.getRange(1, 1, row, 148).getValues();
  let cC = cCSheet.getRange(1, 1, row, 148).getValues();
  let cS = cSSheet.getRange(1, 1, row, 100).getValues();
  let cN = cNSheet.getRange(1, 1, row, 100).getValues();
  let cT = cTSheet.getRange(1, 1, row, 100).getValues();

  let cRn = [];
  let arr = cRnSheet.getRange(1, 1, 1, 148).getDisplayValues();
  if (row > 1) {
    cRn = cRnSheet.getRange(2, 1, row-1, 148).getValues();
    cRn.unshift(arr[0]);
  }
  else { cRn = arr }

  let cP = [];
  arr = cPSheet.getRange(1, 1, 1, 196).getDisplayValues();
  if (row > 1) {
    cP = cPSheet.getRange(2, 1, row-1, 196).getValues();
    cP.unshift(arr[0]);
  }
  else { cP = arr }

  for (let i=1; i<cI.length; i++) {
    cRt[i][2] = '', cRt[i][3] = '', cRt[i][52+tRow-3] = cRt[i][52+tRow-3-1];
    cV[i][2] = '', cV[i][3] = '';
    cL[i][2] = '', cL[i][3] = '';
    cC[i][2] = '', cC[i][3] = '';
    cS[i][2] = '', cS[i][3] = '';
    cN[i][2] = '', cN[i][3] = '';
    cT[i][2] = '', cT[i][3] = '';
  }

  row = cISheet.getLastRow();
  const vRt = vRtSheet.getRange(1, 1, row, 196).getValues();

  let src = Array(iNum+1);

  row = tRow;
  for (let i=0; i<iNum; i++) {
    src[i] = pSheet.getRange(row, 3, 1, rNum).getValues()[0];
    row += nNum;
  }
  src[iNum] = pSheet.getRange(1, 3, 1, rNum).getValues()[0];
  src = src[0].map((_, c) => src.map(r => r[c]));

  src.sort(function(a,b){
    if (a[7] > b[7]) { return 1 }
    else { return -1 }
  })

  for (let i=0; i<rNum; i++) {

    let numV = 1, f = true;
    while (f) {
      if (i+1<rNum) {
        if (src[i+1][7] === src[i][7]) {
          i++;
          numV++;
        }
        else { f = false; }
      }
      else { f = false; }
    }

    row = cI.findIndex(x => x[0] === src[i][7]);
    cI = cInfo(row, src.slice(i-numV+1,i+1), cI);
    cRn = cRank(row, src.slice(i-numV+1,i+1), cRn);
    cRt = cRatio(row, src.slice(i-numV+1,i+1), vRt, cRt);
    cP = cPeriod(row, src.slice(i-numV+1,i+1), cP);
    cV = cOther1(row, src.slice(i-numV+1,i+1), cV, 4);
    cL = cOther1(row, src.slice(i-numV+1,i+1), cL, 5);
    cC = cOther1(row, src.slice(i-numV+1,i+1), cC, 6);
    cS = cOther2(row, src.slice(i-numV+1,i+1), cS, 10);
    cN = cOther2(row, src.slice(i-numV+1,i+1), cN, 11);
    cT = cOther2(row, src.slice(i-numV+1,i+1), cT, 12);
  }

  writeIndicator('c', cI, cRn, cRt, cP, cV, cL, cC, cS,cN, cT);

  pSheet.getRange(tRow, 1).setValue('channel')
}

//  vID, vTitle, vDate, dur, cntV, cntL, cntC, cID, cTitle, cDate,
//  sub, cntN, totV, vDesc, vURL, vTmb, vTags, cDesc, cURL, cTmb,
//  cCustom, rank

function cInfo(row, src, data) {

  let d = Array(8);
  d[0] = src[0][7];
  d[1] = src[0][8];
  d[2] = src.map(x => x[0]).join();
  d[3] = src[0][20];
  d[4] = src[0][9];
  d[5] = src[0][17];
  d[6] = src[0][18];
  d[7] = src[0][19];

  if (row != -1) {
    d[2] = new Set(d[2].split(',').concat(data[row][2].split(',')));
    d[2] = Array.from(d[2]).join();
    data[row] = d;
  }
  else { data.push(d) }

  return data
}

function cRank(row, src, data) {

  if (row != -1) {
    data[row][1] = src[0][8];
    data[row][52+tRow-3] = src.sort(function(a,b){return(a[iNum] - b[iNum]);})[0][iNum];
    //前日との差分
    if (data[row][52+tRow-3-48]) {
      data[row][52+tRow-3+48] = data[row][52+tRow-3] - data[row][52+tRow-3-48]
    }
    //直近1h
    if (data[row][52+tRow-3] < data[row][2]) {
      data[row][2] = data[row][52+tRow-3];
      data[row][3] = rDay + data[0][52+tRow-3];
    }
  }
  else {
    let d = Array(148);
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[52+tRow-3] = src.sort(function(a,b){return(a[iNum] - b[iNum]);})[0][iNum];
    d[2] = d[52+tRow-3];
    d[3] = rDay + data[0][52+tRow-3];
    data.push(d);
  }
  return data
}

function cRatio(row, src, vRt, data) {

  sumR = vRt.filter(function(v){return v[0] === src[0][0];}).reduce((sum, x) => sum + x[148+tRow-3], 0);

  if (row != -1) {
    data[row][1] = src[0][8];
    data[row][148+tRow-3] = sumR;
    data[row][52+tRow-3] = data[row][52+tRow-3-1] + sumR;
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
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[148+tRow-3] = sumR;
    d[52+tRow-3] = sumR;
    data.push(d);
  }
  return data
}

function cPeriod(row, src, data) {

  if (row != -1) {
    data[row][1] = src[0][8];
    //ランクイン開始日時
    if (data[row][148+tRow-3-1]) { data[row][148+tRow-3] = data[row][148+tRow-3-1] }
    else { data[row][148+tRow-3] = rDay + data[0][148+tRow-3] }
    //ランクイン期間
    if (data[row][52+tRow-3-1]) { data[row][52+tRow-3] = data[row][52+tRow-3-1] + 0.5 }
    else { data[row][52+tRow-3] = 0.5 }
    //最長期間の更新
    if (data[row][52+tRow-3] > data[row][2]) {
      data[row][2] = data[row][52+tRow-3];
      data[row][3] = data[row][148+tRow-3] + '～' + rDay + data[0][148+tRow-3];
    }
  }
  else {
    let d = Array(196);
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[2] = 0.5;
    d[3] = rDay + data[0][148+tRow-3] + '～' + rDay + data[0][148+tRow-3];
    d[52+tRow-3] = 0.5;
    d[148+tRow-3] = rDay + data[0][148+tRow-3];
    data.push(d);
  }
  return data
}

function cOther1(row, src, data, clm) {

  if (row != -1) {
    data[row][0] = src[0][7];
    data[row][1] = src[0][8];
    data[row][52+tRow-3] = src.map(x => x[clm]).reduce((sum, x) => sum + x, 0);
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
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[52+tRow-3] = src.map(x => x[clm]).reduce((sum, x) => sum + x, 0);
    data.push(d);
  }
  return data
}

function cOther2(row, src, data, clm) {

  if (row != -1) {
    data[row][0] = src[0][7];
    data[row][1] = src[0][8];
    data[row][52+tRow-3] = src.map(x => x[clm]).reduce((sum, x) => sum + x, 0);
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
    let d = Array(100);
    d[0] = src[0][7];
    d[1] = src[0][8];
    d[52+tRow-3] = src.map(x => x[clm]).reduce((sum, x) => sum + x, 0);
    data.push(d);
  }
  return data
}

//■■■■■■■■■ 指標の記入 ■■■■■■■■■

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
      cSSheet.getRange(1, 1, row, 100).setValues(dS);
      cNSheet.getRange(1, 1, row, 100).setValues(dN);
      cTSheet.getRange(1, 1, row, 100).setValues(dT);
      return;
    default: console.log('■■■■■■■■■■  エラー : writeIndicator  ■■■■■■■■■■')
  }
}

//■■■■■■■■■ 日次、月次 ■■■■■■■■■

function createMonthly() {
  if (tRow != 12 || today != '27' ) { return }

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
  if (tRow != 24) { return }
  if (tomorrow === '01') { return }

  const copySheet = pFile.getSheetByName('dd');
  const newSheet = copySheet.copyTo(pFile);
  newSheet.setName(tomorrow);

  let name = 'g'+ tMonth + tomorrow + '_R' + vCat;
  tmpR.makeCopy(name, cFolder);
}

function transferDaily() {
  if (tRow != 47+3) { return }

  let nFolder = cFolder;
  let name = 'g'+ tMonth + tomorrow + '_R' + vCat;

  if (tomorrow === '01') {
    name = "y" + vCat + '_' + nMonth;
    const nFolderID = pFolder.getFoldersByName(name).next().getId();
    nFolder = DriveApp.getFolderById(nFolderID);
    name = 'g'+ nMonth + tomorrow + '_R' + vCat;
  }

  const nFile_ID = nFolder.getFilesByName(name).next().getId();
  const nFile = SpreadsheetApp.openById(nFile_ID);

  let row = vISheet.getLastRow();
  transferData('vRn', row, nFile);
  transferData('vV', row, nFile);
  transferData('vL', row, nFile);
  transferData('vC', row, nFile);

  row = cISheet.getLastRow();
  transferData('cRn', row, nFile);
  transferData('cV', row, nFile);
  transferData('cL', row, nFile);
  transferData('cC', row, nFile);
  transferData('cN', row, nFile);
  transferData('cS', row, nFile);
  transferData('cT', row, nFile);

  transferData1('v', nFile);
  transferData1('c', nFile);
}

function transferData(sName, row, nFile) {

  const cSheet = rFile.getSheetByName(sName);
  const nSheet = nFile.getSheetByName(sName);

  let src = cSheet.getRange(1, 1, row, 100).getValues();
  src = src.filter(function(x){return x[99] != '';});
  const src1 = src.map(x => x.slice(52));
  src = src.map(x => x.slice(0, 4));

  nSheet.getRange(1, 1, src.length, 4).setValues(src);
  nSheet.getRange(1, 5, src.length, 48).setValues(src1);
}

function transferData1(cat, nFile) {

  const oISheet = rFile.getSheetByName(cat+'I')
  let row = oISheet.getLastRow();

  const oRtSheet = rFile.getSheetByName(cat+'Rt');
  const oPSheet = rFile.getSheetByName(cat+'P');
  const nISheet = nFile.getSheetByName(cat+'I');
  const nRtSheet = nFile.getSheetByName(cat+'Rt');
  const nPSheet = nFile.getSheetByName(cat+'P');

  let src = oISheet.getRange(1, 1, row, 4).getValues();
  let src1 = oRtSheet.getRange(1, 1, row, 196).getValues();

  for (let i=0; i<row; i++) { src1[i] = src1[i].concat(src[i]) }

  src1 = src1.filter(function(x){return x[195] != '';});

  let src2 = src1.map(x => x.slice(52, 52+48));
  let src3 = src1.map(x => x.slice(148, 148+48));
  src = src1.map(x => x.slice(196));
  src1 = src1.map(x => x.slice(0, 4));

  row = src.length;
  nISheet.getRange(1, 1, row, 4).setValues(src);
  nRtSheet.getRange(1, 1, row, 4).setValues(src1);
  nRtSheet.getRange(1, 5, row, 48).setValues(src2);
  nRtSheet.getRange(1, 101, row, 48).setValues(src3);

  src1 = oPSheet.getRange(1, 1, row, 196).getValues();
  src1 = src1.filter(function(x){return x[195] != '';});
  src2 = src1.map(x => x.slice(52, 52+48));
  src3 = src1.map(x => x.slice(148));
  src1 = src1.map(x => x.slice(0, 4));

  row = src.length;
  nPSheet.getRange(1, 1, row, 4).setValues(src1);
  nPSheet.getRange(1, 5, row, 48).setValues(src2);
  nPSheet.getRange(1, 101, row, 48).setValues(src3);
}