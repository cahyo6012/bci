function main() {
  let cmd
  try {
    cmd = require('./lib/' + process.argv[2])
  } catch (error) {
    console.log(`Perintah ${ process.argv[2] } Tidak Diketahui...`)
    process.exit()
  }
  cmd()
}

main()