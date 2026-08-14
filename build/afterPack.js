const { execFileSync } = require('child_process')

/** Apple Silicon refuses unsigned apps with “damaged and can’t be opened”. */
exports.default = async function afterPack(context) {
  if (context.electronPlatformName !== 'darwin') return
  const appPath = `${context.appOutDir}/${context.packager.appInfo.productFilename}.app`
  execFileSync('codesign', ['--force', '--deep', '--sign', '-', appPath], {
    stdio: 'inherit'
  })
}
