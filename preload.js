/*
 * @Author: Libra
 * @Date: 2022-06-28 17:26:18
 * @LastEditTime: 2022-06-28 17:37:10
 * @LastEditors: Libra
 * @Description: 
 * @FilePath: /word_to_excel/preload.js
 */
window.addEventListener('DOMContentLoaded', () => {
  const replaceText = (selector, text) => {
    const element = document.getElementById(selector)
    if (element) element.innerText = text
  }

  for (const dependency of ['chrome', 'node', 'electron']) {
    replaceText(`${dependency}-version`, process.versions[dependency])
  }
})