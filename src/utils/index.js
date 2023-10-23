import * as XLSX from 'xlsx/xlsx.mjs'
import XLSXStyle from 'xlsx-style-fixedver'
/**
 * 将 String 转换成 ArrayBuffer
 * @method 类型转换
 * @param {String} [s] wordBook内容
 * @return {Array} 二进制流数组
 */
function s2ab(s) {
  let buf = null

  if (typeof ArrayBuffer !== 'undefined') {
    buf = new ArrayBuffer(s.length)
    const view = new Uint8Array(buf)

    for (let i = 0; i !== s.length; ++i) {
      view[i] = s.charCodeAt(i) & 0xFF
    }

    return buf
  }

  buf = new Array(s.length)

  for (let i = 0; i !== s.length; ++i) {
    // 转换成二进制流
    buf[i] = s.charCodeAt(i) & 0xFF
  }

  return buf
}

/**
 * 方案一：利用 URL.createObjectURL 下载 （以下选用）
 * 方案二：通过 file-saver 插件实现文件下载
 * @method 文件下载
 * @param {Object} [obj] 导出内容 Blob 对象
 * @param {String} [fileName] 文件名 下载是生成的文件名
 * @return {void}
 */
function saveAs(obj, fileName) {
  const aLink = document.createElement('a')

  // eslint-disable-next-line eqeqeq
  if (typeof obj == 'object' && obj instanceof Blob) {
    aLink.href = URL.createObjectURL(obj) // 创建blob地址
  }

  aLink.download = fileName
  aLink.click()
  setTimeout(function() {
    URL.revokeObjectURL(obj)
  }, 100)
}

/**
 * @method 数据导出excel
 * @param {Object} [data] 工作表数据内容
 * @param {String} [name] 导出excel文件名
 * @param {Number} [merges] 表头合并列数
 * @param {Boolean} [save] 直接下载或返回bolb文件
 */
export function exportExcel(data, name, merges, save = true) {
  return new Promise((resolve) => {
    let index = 0
    // 合并单元格 s:开始位置 e:结束位置 r:行 c:列
    const datamerges = [
      // 实际情况根据业务需求进行
      { s: { c: 0, r: 0 }, e: { c: merges, r: 0 }},
      { s: { c: 1, r: data.length + 1 }, e: { c: 3, r: data.length + 1 }}
    ]
    const worksheet1 = XLSX.utils.json_to_sheet(data, { origin: 'A2' }) // origin:指定某一行开始导入表格数据
    const itemWidth = []
    const itemHeight = []
    worksheet1['!merges'] = datamerges
    worksheet1.A1 = {
      t: 's',
      v: name
    }
    const total = `B${data.length + 2}`
    worksheet1[total] = {
      t: 's',
      v: '合计'
    }
    for (const key in worksheet1) {
      index++
      // 前两行高度为45，其后行高为25
      if (index <= 2) {
        itemHeight.push({ hch: 45 })
      } else {
        itemHeight.push({ hch: 25 })
      }
      // 所有单元格居中
      if (key !== '!cols' && key !== '!merges' && key !== '!ref' && key !== '!rows') {
        worksheet1[key].s = {
          alignment: {
            horizontal: 'center',
            vertical: 'center'
          }
        }
      }
      // 所有单元格列宽为15
      itemWidth.push({ wch: 15 })
      // A1单元格加粗居中
      if (key === 'A1') {
        worksheet1[key].s = {
          font: {
            bold: true,
            sz: 20
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center'
          }
        }
      }
      if (key === total) {
        worksheet1[key].s = {
          font: {
            bold: true
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center'
          }
        }
      }
      // 表头加粗居中
      if (key.replace(/[^0-9]/ig, '') === '2') {
        worksheet1[key].s = {
          font: {
            bold: true
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center'
          }
        }
      }
    }
    worksheet1['!cols'] = itemWidth // 列宽
    worksheet1['!rows'] = itemHeight // 行高
    // return
    const sheetNames = Object.keys({ '总表': worksheet1 })
    const workbook = {
      SheetNames: sheetNames, // 保存的工作表名
      Sheets: { '总表': worksheet1 } // 与表名对应的表数据
    }
    // // excel的配置项
    const wopts = {
      bookType: 'xlsx', // 生成的文件类型
      bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
      type: 'binary'
    }
    // attempts to write the workbook
    // const wbout = styleXLSX.write(workbook, wopts)
    const wbout = XLSXStyle.write(workbook, wopts)
    try {
      wbout.then(res => {
        const wbBlob = new Blob([s2ab(res)], {
          type: 'application/octet-stream'
        })
        if (save) {
          saveAs(wbBlob, name + '.' + 'xlsx')
        } else {
          resolve(wbBlob)
        }
      })
    } catch (error) {
      const wbBlob = new Blob([s2ab(wbout)], {
        type: 'application/octet-stream'
      })
      if (save) {
        saveAs(wbBlob, name + '.' + 'xlsx')
      } else {
        resolve(wbBlob)
      }
    }
  })
}
