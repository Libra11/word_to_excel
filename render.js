/*
 * @Author: Libra
 * @Date: 2022-06-28 17:39:41
 * @LastEditTime: 2022-07-01 11:11:52
 * @LastEditors: Libra
 * @Description: 渲染进程
 * @FilePath: /word_to_excel/render.js
 */
// 获取 textarea
const textarea = document.getElementById("textarea");
const textarea2 = document.getElementById("textarea2");
const input = document.getElementById("input");
// 获取 select dom
const select = document.getElementById("select");
const select2 = document.getElementById("select2");
const select3 = document.getElementById("select3");
const select4 = document.getElementById("select4");
// 获取处理 button
const handleButton = document.getElementById("btn");
const checkButton = document.getElementById("btn2");
const downloadButton = document.getElementById("btn3");
const clearButton = document.getElementById("btn4");
const oriCheckButton = document.getElementById("btn5");
const transButton = document.getElementById("button");
const trueFalseButton = document.getElementById("btn6");

const err = document.getElementById("err");

let originText = "";
let option = {
  selectedText: "pause",
  selectedText3: "pause",
  selectedText2: "&",
  selectedText4: "$",
};
let fileName = "";
// 监听文本框的变化
textarea.addEventListener("input", () => {
  originText = textarea.value;
});
// 监听 input 的变化
input.addEventListener("input", () => {
  fileName = input.value;
});
// 监听 select 的变化
select.addEventListener("change", () => {
  option.selectedText = select.options[select.selectedIndex].value;
});
select2.addEventListener("change", () => {
  option.selectedText2 = select2.options[select2.selectedIndex].text;
});
select3.addEventListener("change", () => {
  option.selectedText3 = select3.options[select3.selectedIndex].value;
});
select4.addEventListener("change", () => {
  option.selectedText4 = select4.options[select4.selectedIndex].text;
});
// 监听 button 的点击
handleButton.addEventListener("click", () => {
  handleText();
});
checkButton.addEventListener("click", () => {
  check();
});
downloadButton.addEventListener("click", () => {
  downloadText();
});
clearButton.addEventListener("click", () => {
  clearText();
});
oriCheckButton.addEventListener("click", () => {
  checkOriginString();
});
transButton.addEventListener("click", () => {
  transText();
});
trueFalseButton.addEventListener("click", () => {
  genTrueFalseQues();
});

// 下载文本
function downloadText() {
  const regex = getRegex();
  const regex2 = getRegex2();
  replaceRegex(regex, option.selectedText2, true)
    .replaceRegex(regex2, option.selectedText4)
    .replaceRegex(wordEnterRegex, "")
    .trim()
    .trimFirst();
  const arrData = splitData(originText);
  downloadExcel(arrData);
}
function genTrueFalseQues() {
  const regex = getRegex();
  replaceRegex(regex, option.selectedText2, true)
    .replaceRegex(wordEnterRegex, "")
    .trim()
    .trimFirst();
  const arrData = splitSelect(originText);
  downloadExcel(arrData);
}
// 检测
function check() {
  const regex = getRegex();
  const regex2 = getRegex2();
  replaceRegex(regex, option.selectedText2, true)
    .replaceRegex(regex2, option.selectedText4)
    .replaceRegex(wordEnterRegex, "")
    .trim()
    .trimFirst();
  const arrData = splitData(originText);
}

// 原始字符检测
function checkOriginString() {
  if (checkString(originText, option.selectedText2)) {
    err.innerText =
      "试题中检测到" + option.selectedText2 + "字符, 请更换题号分隔符";
    return;
  }
  if (checkString(originText, option.selectedText4)) {
    err.innerText =
      "试题中检测到" + option.selectedText4 + "字符, 请更换选项分隔符";
    return;
  }
  err.innerText = "检测通过！！";
}

// 转换两个 textarea 的内容
function transText() {
  textarea.value = textarea2.value;
  textarea2.value = "";
}

// 处理文本
function handleText() {
  const regex = getRegex();
  const regex2 = getRegex2();
  replaceRegex(regex, option.selectedText2, true)
    .replaceRegex(regex2, option.selectedText4)
    .replaceRegex(wordEnterRegex, "")
    .trim()
    .trimFirst();
  textarea2.value = originText;
}

function getRegex() {
  switch (option.selectedText) {
    case "pause":
      return numberPauseRegex;
    case "dot":
      return numberDotRegex;
    case "chineseDot":
      return numberChineseDotRegex;
    case "blank":
      return numberSpaceRegex;
    default:
      return numberPauseRegex;
  }
}

function getRegex2() {
  switch (option.selectedText3) {
    case "pause":
      return lashPauseRegex;
    case "dot":
      return lashDotRegex;
    case "chineseDot":
      return lashChineseDotRegex;
    case "blank":
      return lashSpaceRegex;
    default:
      return lashPauseRegex;
  }
}

function clearText() {
  textarea2.value = "";
  textarea.value = "";
}

/**
 * 正则表达式
 */
// 匹配 数字+顿号
const numberPauseRegex = /\d+、/g;
// 匹配 数字+点+非数字
const numberDotRegex = /\d+\.\D/g;
// const numberDotRegex = /\d+\./g;
// 匹配 数字+中文点
const numberChineseDotRegex = /\d+\．/g;
// 匹配 数字+空格
const numberSpaceRegex = /\d+ /g;
// 匹配 A-D+顿号
const lashPauseRegex = /[A-D]+、/g;
// 匹配 A-D+点
const lashDotRegex = /[A-D]+\./g;
// 匹配 A-D+中文点
const lashChineseDotRegex = /[A-D]+\．/g;
// 匹配 A-D+空格
const lashSpaceRegex = /[A-D]+ /g;
// 匹配 word 的段落标记
const wordEnterRegex = /\n/g;

/**
 * 工具函数
 */
// 替换正则表达式内容为特定字符
function replaceRegex(regex, replace, isFirst = false) {
  originText = originText.replace(regex, (x) => {
    if (isFirst) {
      x = replace + x[x.length - 1];
    } else {
      x = replace;
    }
    return x;
  });
  return this;
}
// 去掉所有空格
function trim() {
  originText = originText.replace(/\s+/g, "");
  return this;
}
// 去掉第一个字符
function trimFirst() {
  if (originText[0] !== option.selectedText2) return this;
  originText = originText.substring(1);
  return this;
}
// 将分割后的数据转为二维数组
function splitData(data) {
  const dataArray = data.split(option.selectedText2);
  const a = +dataArray.length;
  err.innerText = "共有" + a + "道题";
  return splitArray(dataArray, option.selectedText4);
}
// 将数组里的每一项通过指定的分割符分割成二维数组
function splitArray(array, symbol) {
  const arr = [];
  for (let i = 0; i < array.length; i++) {
    const arr2 = array[i].split(symbol);
    arr.push(arr2);
  }
  return arr;
}
// 选择题
function splitSelect(data) {
  const dataArray = data.split(option.selectedText2);
  const a = +dataArray.length;
  err.innerText = "共有" + a + "道题";
  return splitArraySelect(dataArray);
}
function splitArraySelect(array) {
  const arr = [];
  for (let i = 0; i < array.length; i++) {
    const arr2 = [array[i], "正确", "错误"];
    arr.push(arr2);
  }
  return arr;
}
// 将数组转为excel文件并下载
function downloadExcel(data) {
  if (fileName === "") {
    err.innerText = "请输入文件名!!";
    return;
  }
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  XLSX.writeFile(workbook, fileName + ".xlsx");
}
// 检测字符串中是否含有某个字符
function checkString(str, symbol) {
  return str.indexOf(symbol) !== -1;
}
