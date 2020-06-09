const path = require('path')
const XLSX = require('xlsx')
const fs = require('fs')

const RESULT_PATH = path.resolve(__dirname, '..', 'results', process.argv[3])
const LIST_KPP = ['031', '032', '033', '034', '035', '036', '037', '039', '085', '086']

function countFirms(firms) {
  const firmIds = []
  for (let firm of firms) firmIds.push(firm.FIRMID)
  const uniqueFirmIds = new Set(firmIds)
  return uniqueFirmIds.size
}

function countRegisteredFirms(mfwp) {
  const firmIds = []
  for (let m of mfwp) firmIds.push(m.id)
  const uniqueFirmIds = new Set(firmIds)
  return uniqueFirmIds.size
}

function countRegisteredFirmsInKpp(mfwp = [], kpp) {
  const filteredFirms = mfwp.filter(v => v.kode_kpp == kpp)
  const firmIds = []
  for (let fm of filteredFirms) firmIds.push(fm.id)
  const uniqueFirmIds = new Set(firmIds)
  return uniqueFirmIds.size
}

function generateRekapCell(data, tahun = new Date().getFullYear(), ) {
  const ws = require('./template.json')
  
  ws.A1.v = ws.A1.h = ws.A1.w = `Rekap Data Proyek Konstruksi Tahun ${ tahun }`
  ws.A1.r = `<t>${ ws.A1.v }</t>`
  for (let i = 6; i < 16; i++){
    ws[`D${i}`].v = data.jumlahProyek[i - 6]
    ws[`E${i}`].v = data.jumlahWpBelumTerdaftar[i - 6]
    ws[`F${i}`].v = data.jumlahWpTerdaftarDalamKpp[i - 6]
    ws[`G${i}`].v = data.jumlahWpTerdaftarLuarKpp[i - 6]
    ws[`H${i}`].v = data.jumlahWpTerdaftar[i - 6]
    ws[`I${i}`].v = data.jumlahPerusahaan[i - 6]
  }

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'REKAP')
  XLSX.writeFile(wb, path.resolve(RESULT_PATH, `REKAP ${tahun}.xlsx`), { type: 'binary' })
}

function main() {
  const data = {
    jumlahProyek: [],
    jumlahPerusahaan: [],
    jumlahWpTerdaftar: [],
    jumlahWpBelumTerdaftar: [],
    jumlahWpTerdaftarDalamKpp: [],
    jumlahWpTerdaftarLuarKpp: [],
  }
  
  for (let i = 0; i < LIST_KPP.length; i++) {
    const wb = XLSX.readFile(path.resolve(RESULT_PATH, LIST_KPP[i] + '.xlsx'))
    const firms = XLSX.utils.sheet_to_json(wb.Sheets['FIRMS'])
    const projects = XLSX.utils.sheet_to_json(wb.Sheets['PROJECTS'])
    const mfwp = XLSX.utils.sheet_to_json(wb.Sheets['MFWP'])

    data.jumlahProyek[i] = projects.length
    data.jumlahPerusahaan[i] = countFirms(firms)
    data.jumlahWpTerdaftar[i] = countRegisteredFirms(mfwp)
    data.jumlahWpBelumTerdaftar[i] = data.jumlahPerusahaan[i] - data.jumlahWpTerdaftar[i]
    data.jumlahWpTerdaftarDalamKpp[i] = countRegisteredFirmsInKpp(mfwp, LIST_KPP[i])
    data.jumlahWpTerdaftarLuarKpp[i] = data.jumlahWpTerdaftar[i] - data.jumlahWpTerdaftarDalamKpp[i]
  }
  generateRekapCell(data, process.argv[3])
}

module.exports = main