/*
 * @Author: Libra
 * @Date: 2022-06-28 17:39:41
 * @LastEditTime: 2022-06-29 10:29:45
 * @LastEditors: Libra
 * @Description: 渲染进程
 * @FilePath: /word_to_excel/render.js
 */
// 获取 textarea
const textarea = document.getElementById("textarea");
const textarea2 = document.getElementById("textarea2");
// 获取 select dom
const select = document.getElementById("select");
const select2 = document.getElementById("select2");
const select3 = document.getElementById("select3");
// 获取处理 button
const handleButton = document.getElementById("btn");
const checkButton = document.getElementById("btn2");
const downloadButton = document.getElementById("btn3");
const clearButton = document.getElementById("btn4");

let originText = "";
let option = {
  selectedText: "pause",
  selectedText2: "&",
  length: "4",
};
// 监听文本框的变化
textarea.addEventListener("input", () => {
  originText = textarea.value;
});
// 监听 select 的变化
select.addEventListener("change", () => {
  option.selectedText = select.options[select.selectedIndex].value;
});
select2.addEventListener("change", () => {
  option.selectedText2 = select2.options[select2.selectedIndex].text;
});
select3.addEventListener("change", () => {
  option.length = select3.options[select3.selectedIndex].value;
});
// 监听 button 的点击
handleButton.addEventListener("click", () => {
  handleText();
});
checkButton.addEventListener("click", () => {
  check();
});
downloadButton.addEventListener("click", () => {
  handleText();
});
clearButton.addEventListener("click", () => {
  clearText();
});

// 下载文本
function downloadText() {
  const regex =
    option.selectedText === "pause" ? numberPauseRegex : numberDotRegex;
  const regex2 =
    option.selectedText === "pause" ? lashPauseRegex : lashDotRegex;
  replaceRegex(regex, option.selectedText2)
    .replaceRegex(regex2, option.selectedText2)
    .replaceRegex(wordEnterRegex, "")
    .trim()
    .trimFirst();
  const arrData = splitData(originText);
  downloadExcel(arrData);
}
// 检测
function check() {
  const regex =
    option.selectedText === "pause" ? numberPauseRegex : numberDotRegex;
  const regex2 =
    option.selectedText === "pause" ? lashPauseRegex : lashDotRegex;
  replaceRegex(regex, option.selectedText2)
    .replaceRegex(regex2, option.selectedText2)
    .replaceRegex(wordEnterRegex, "")
    .trim()
    .trimFirst();
  const arrData = splitData(originText);
}

// 处理文本
function handleText() {
  const regex =
    option.selectedText === "pause" ? numberPauseRegex : numberDotRegex;
  const regex2 =
    option.selectedText === "pause" ? lashPauseRegex : lashDotRegex;
  replaceRegex(regex, option.selectedText2)
    .replaceRegex(regex2, option.selectedText2)
    .replaceRegex(wordEnterRegex, "")
    .trim()
    .trimFirst();
  textarea2.value = originText;
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
// 匹配 数字+点
const numberDotRegex = /\d+\./g;
// 匹配 A-D+顿号
const lashPauseRegex = /[A-D]+、/g;
// 匹配 A-D+点
const lashDotRegex = /[A-D]+\./g;
// 匹配 word 的段落标记
const wordEnterRegex = /\n/g;

/**
 * 工具函数
 */
// 替换正则表达式内容为特定字符
function replaceRegex(regex, replace) {
  originText = originText.replace(regex, replace);
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
  const b = Number(option.length) + 1;
  alert("检测到" + a / b + "道题");
  return splitArray(dataArray, b);
}
// 一维数组每 num 个拼接为一个二维数组
function splitArray(array, num) {
  const result = [];
  for (let i = 0; i < array.length; i += num) {
    result.push(array.slice(i, i + num));
  }
  return result;
}
// 将数组转为excel文件并下载
function downloadExcel(data) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  XLSX.writeFile(workbook, "out.xlsx");
}
