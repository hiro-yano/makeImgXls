// setTimeoutとclearTimeoutを定義しないとエラーになるため、書いておきます
global.setTimeout = function (fn) { fn() }
global.clearTimeout = function () { }

const { Artboard, Text } = require('scenegraph')
const Vue = require('vue').default
const App = require('./App.vue').default

function createDialog (id = 'dialog') {
  document.body.innerHTML = `<dialog id="${id}"><div id="container"></div></dialog>`
  let dialog = document.getElementById(id)
  new Vue({
    el: '#container',
    components: { App },
    render (h) {
      return h(App, { props: { dialog } })
    }
  })

  return dialog
}

function replaceText (children, prev, after) {
  children.forEach(elm => {
    if (elm instanceof Text) {
      elm.text = elm.text.replace(prev, after)
    }
  })
}
// エントリポイントとなるメソッドです
async function myPluginCommand (selection, root) {
  const dialog = createDialog()
  const result = await dialog.showModal()
  if (result) {
    root.children.forEach(elm => {
      if (elm instanceof Artboard) {
        replaceText(elm.children, new RegExp(result[0], 'g'), result[1])
      }
    })
  }
}
// manifest.jsonで指定したコマンドIDとファンクションを紐づけます
module.exports = {
  commands: {
    myPluginCommand: myPluginCommand
  }
}
