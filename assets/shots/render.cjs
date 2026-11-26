// Renders every page of this folder to a PNG beside it, in a window that is
// never shown: the pictures are drawn, and nothing is taken off the screen.
// Run with any Electron (this repository has none of its own):
//
//   <path to>/electron assets/shots/render.cjs
//
// A page is 1500 x 850 (fixed in shots.css). A window cannot be larger than
// the screen, so the window is 938 x 532 with the page zoomed out to fit, at
// twice the density: the picture comes out 1876 x 1064 on any screen of
// 1920 x 1200 or more. Measured on the way here: headless Edge wrote no file
// and said nothing, an offscreen window cannot be captured («UnknownVizError»),
// and a capture through the debugger never returns from a hidden window.
const { app, BrowserWindow } = require('electron')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')

app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'syw-shots-')))
app.commandLine.appendSwitch('force-device-scale-factor', '2')
// The app stays until every page is done
app.on('window-all-closed', () => {})

app.whenReady().then(async () => {
  for (const name of fs.readdirSync(__dirname).filter(one => one.endsWith('.html'))) {
    const win = new BrowserWindow({
      show: false,
      width: 938,
      height: 532,
      useContentSize: true,
      frame: false,
      skipTaskbar: true,
      focusable: false,
      paintWhenInitiallyHidden: true,
      webPreferences: { zoomFactor: 0.625 },
    })

    await win.loadFile(path.join(__dirname, name))
    // A moment more, so the fonts are in
    await new Promise(done => setTimeout(done, 600))

    const image = await win.webContents.capturePage()
    const out = path.join(__dirname, name.replace(/\.html$/, '.png'))

    fs.writeFileSync(out, image.toPNG())
    console.log(`${name} -> ${path.basename(out)} ${image.getSize().width} x ${image.getSize().height}`)
    win.destroy()
  }

  app.quit()
})
