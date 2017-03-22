const ipc = require('electron').ipcRenderer

const selectDirBtn = document.getElementById('oef')

const filterBtn = document.getElementById('filterBtn')

const excelBtn = document.getElementById('excelBtn')

const saveBtn = document.getElementById('saveBtn')

const symbol = document.getElementById('symbol')

const times = document.getElementById('times')

filterBtn.disabled = true
saveBtn.disabled = true
excelBtn.disabled = true
times.disabled = true

selectDirBtn.addEventListener('click', function (event) {
    document.getElementById("overlay").className = "overlay overlay-hugeinc"
    ipc.send('open-file-dialog')
})

ipc.on('selected-directory', (event, path) => {
    document.getElementById('selected-file').innerHTML = `You selected: ${path}`
    document.getElementById("overlay").className = "overlay hidden-overlay-hugeinc"
})
let r = ''
ipc.on('result', (event, result) => {
    document.getElementById('file-result').innerHTML = `${result}`
    if (result != null) {
        r = result
        filterBtn.disabled = false
        saveBtn.disabled = false
        excelBtn.disabled = false
        times.disabled = false
        document.getElementById("overlay").className = "overlay hidden-overlay-hugeinc"
    }
})

ipc.on('dismissed', (event, dismissed) => {
    if (dismissed != null) {
        document.getElementById("overlay").className = "overlay hidden-overlay-hugeinc"
    }
})

filterBtn.addEventListener('click', (event) => {
    if (times.value.trim() == '') {
        alert('Please input a number')
    } else {
        document.getElementById("overlay").className = "overlay overlay-hugeinc"
        ipc.send('click-filterBtn', symbol.value, times.value)
    }
})

saveBtn.addEventListener('click', (event) => {
    document.getElementById("overlay").className = "overlay overlay-hugeinc"
    ipc.send('save-file-dialog', r)
})

excelBtn.addEventListener('click', (event) => {
    document.getElementById("overlay").className = "overlay overlay-hugeinc"
    ipc.send('save-excel-dialog', r)
})