const path = require('path')
const XLSX = require('xlsx')

const RESULT_PATH = path.resolve(__dirname, '..', 'results', process.argv[3])
const DATA_PATH = path.resolve(RESULT_PATH, process.argv[4])
const LIST_KPP = ['031', '032', '033', '034', '035', '036', '037', '039', '085', '086']

function main() {
  const wb = XLSX.readFile(DATA_PATH)
  const firms = XLSX.utils.sheet_to_json(wb.Sheets['FIRMS'])
  const projects = XLSX.utils.sheet_to_json(wb.Sheets['PROJECTS'])
  const pos = XLSX.utils.sheet_to_json(wb.Sheets['KD_POS'])
  const mfwp = XLSX.utils.sheet_to_json(wb.Sheets['MFWP'])

  const list_kpp = []
  for (let firm of firms) {
    let kpp = ''
    for (p of pos) {
      if (firm.FIRM_POSTCODE == p['Kode POS']) kpp = p.KPPADM
    }
    list_kpp.push(kpp)
  }
  for(let i = 0; i < firms.length; i++) firms[i].KPPADM_BY_POSTCODE = list_kpp[i]

  const firmsByKpp = []
  const firmIdsByKpp = []
  const projectIdsByKpp = []
  for (let firm of firms) {
    const idx = LIST_KPP.indexOf(firm.KPPADM_BY_POSTCODE)
    if (idx >= 0) {
      if(!firmsByKpp[idx]) firmsByKpp[idx] = []
      firmsByKpp[idx].push(firm)
      
      if(!firmIdsByKpp[idx]) firmIdsByKpp[idx] = []
      firmIdsByKpp[idx].push(firm.FIRMID)
      
      if(!projectIdsByKpp[idx]) projectIdsByKpp[idx] = []
      projectIdsByKpp[idx].push(firm.PROJECTID)
    }
  }

  const projectsByKpp = []
  for (let project of projects) {
    for (let i = 0; i < projectIdsByKpp.length; i++) {
      if (projectIdsByKpp[i].includes(project.PROJECTID)) {
        if(!projectsByKpp[i]) projectsByKpp[i] = []
        projectsByKpp[i].push(project)
      }
    }
  }

  const mfwpsByKpp = []
  for (let m of mfwp) {
    for (let i = 0; i < firmIdsByKpp.length; i++) {
      if (firmIdsByKpp[i].includes(m.id)) {
        if(!mfwpsByKpp[i]) mfwpsByKpp[i] = []
        mfwpsByKpp[i].push(m)
      }
    }
  }

  for (let i = 0; i < LIST_KPP.length; i++) {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(firmsByKpp[i]), 'FIRMS')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(projectsByKpp[i]), 'PROJECTS')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mfwpsByKpp[i]), 'MFWP')
    XLSX.writeFile(wb, path.resolve(RESULT_PATH, LIST_KPP[i] + '.xlsx'), { type: 'binary' })
  }
}

module.exports = main