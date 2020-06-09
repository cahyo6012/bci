require('dotenv').config()
const path = require('path')
const XLSX = require('xlsx')
const cheerio = require('cheerio')
const cliProgress = require('cli-progress')
const request = require('request').defaults({
  rejectUnauthorized: false,
  followAllRedirects: false,
  jar: true,
  headers: {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36',
  }
})

const { BASE_URL, USER_SIKKA, PASS_SIKKA } = process.env
const DATA_PATH = path.resolve(__dirname, '..', 'assets')
const RESULT_PATH = path.resolve(__dirname, '..', 'results')

function login(username, password) {
  return new Promise((resolve, reject) => {
    console.log('Mencoba Login Appportal...')
    const url = BASE_URL + 'login/login/loging_simpel'
    request.post(url, {
      form: { username, password, sublogin: 'Login' }
    }, (err, res) => {
      if (err) reject(err)
      if (!!res.headers.refresh) reject(new Error('Login Appportal Gagal. Username atau Password Salah...'))
      console.log('Login Appportal Berhasil...')
      resolve(true)
    })
  })
}

function getDistinctFirmsId(firms) {
  const distinctFirmsId = []
  for (let firm of firms) {
    if(!distinctFirmsId.includes(firm.FIRMID)) distinctFirmsId.push(firm.FIRMID)
  }
  return distinctFirmsId
}

function getDistinctFirms(ids, firms) {
  const distinctFirms = []
  for (let id of ids) {
    const firm = firms.find(f => f.FIRMID == id)
    let firmName = firm.FIRM_NAME
    
    const chars = ['PT', 'CV', '(', '-']
    for (char of chars) {
      const idx = firmName.indexOf(char)
      if(idx != -1) firmName = firmName.substring(0, idx).trim()
    }
    distinctFirms.push({ id, name: firmName })
  }
  return distinctFirms
}

function getListMfwp(firms) {
  return new Promise(async (resolve, reject) => {
    const bar = new cliProgress.SingleBar({
      format: 'Mengambil Data MFWP [{bar}] {percentage}% | ETA: {eta}s | {value}/{total}',
    })
    bar.start(firms.length, 0)
    
    const listMfwp = []
    for (let i = 0; i < firms.length; i++) {
      const mfwp = await getMfwp(firms[i])
      listMfwp.push(...mfwp)
      bar.increment()
    }

    bar.stop()
    resolve(listMfwp)
  })
}

function getMfwp(firm) {
  return new Promise((resolve, reject) => {
    const url = BASE_URL + 'portal/masterfile/hasil.php'
    try {
      request.post(url, {
        form: { input: firm.name, kriteria1: 2, kriteria2: 1 }
      }, (err, res) => {
        if (err) resolve([])
        const $ = cheerio.load(res.body)
        const rows = $('tr')
  
        const listMfwp = []
        for (let i = 1; i < rows.length; i++) {
          const cols = $('td', $(rows.get(i)))
  
          const npwp = $(cols[1]).text().replace(/-|\./g, '').trim()
          const nama = $(cols[2]).text().trim()
          const alamat = $(cols[3]).text().trim()
          const tgl_lahir = $(cols[4]).text().trim()
          const no_id = $(cols[5]).text().trim()
          const klu = $(cols[6]).text().trim()
          const { groups: { nip_ar, nama_ar } } = $(cols[7]).text().trim().match(/(?<nip_ar>\w{1,9})\s(?<nama_ar>.*)/)
          const status = $(cols[8]).text().trim()
          const tgl_update = $(cols[9]).text().trim()
          const { groups: { kode_kpp, nama_kpp } } = {} = $(cols[10]).text().trim().match(/(?<kode_kpp>\d{3})\s-\s(?<nama_kpp>.*)/)
          const merk = $(cols[11]).text().trim()
  
          const mfwp = { id: firm.id, npwp, nama, alamat, tgl_lahir, no_id, klu, nip_ar, nama_ar, status, tgl_update, kode_kpp, nama_kpp, merk }
          listMfwp.push(mfwp)
        }
        
        resolve(listMfwp)
      })
    } catch (error) {
      resolve(getMfwp(firm))
    }
  })
}

async function main() {
  try {
    await login(USER_SIKKA, PASS_SIKKA)
  } catch (err) {
    console.log(err)
    process.exit()
  }

  const wb = XLSX.readFile(path.resolve(DATA_PATH, process.argv[3]))
  const firms = XLSX.utils.sheet_to_json(wb.Sheets['FIRMS'])

  const distinctFirmsId = getDistinctFirmsId(firms)
  const distinctFirms = getDistinctFirms(distinctFirmsId, firms)
  const mfwp = await getListMfwp(distinctFirms)

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mfwp), 'MFWP')
  XLSX.writeFile(wb, path.resolve(RESULT_PATH, process.argv[4]), {
    type: 'binary'
  })
}

console.log(__dirname)

module.exports = main