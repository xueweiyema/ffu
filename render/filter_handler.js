const ipc = require('electron').ipcRenderer

const selectDirBtn = document.getElementById('oef')

const filterBtn = document.getElementById('filterBtn')

const symbol = document.getElementById('symbol')

const times = document.getElementById('times')

filterBtn.disabled = true
times.disabled = true

selectDirBtn.addEventListener('click', function (event) {
    ipc.send('open-file-dialog')
})

ipc.on('selected-directory', (event, path) => {
    document.getElementById('selected-file').innerHTML = `You selected: ${path}`
})

ipc.on('result', (event, result) => {
    document.getElementById('file-result').innerHTML = `${result}`
    if (result != null) {
        filterBtn.disabled = false
        times.disabled = false
    }
})

filterBtn.addEventListener('click', (event) => {
    ipc.send('click-filterBtn', symbol.value, times.value)
})