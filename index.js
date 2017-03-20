'use strict'
const electron = require('electron')
const xlsx = require('xlsx')
const os = require('os')
const url = require('url')
const path = require('path')
const {
    app,
    BrowserWindow,
    ipcMain,
    dialog,
    Menu
} = electron

let mainWindow

let template = [{
    label: 'Edit',
    submenu: [{
        label: 'Undo',
        accelerator: 'CmdOrCtrl+Z',
        role: 'undo'
    }, {
        label: 'Redo',
        accelerator: 'Shift+CmdOrCtrl+Z',
        role: 'redo'
    }, {
        type: 'separator'
    }, {
        label: 'Cut',
        accelerator: 'CmdOrCtrl+X',
        role: 'cut'
    }, {
        label: 'Copy',
        accelerator: 'CmdOrCtrl+C',
        role: 'copy'
    }, {
        label: 'Paste',
        accelerator: 'CmdOrCtrl+V',
        role: 'paste'
    }, {
        label: 'Select All',
        accelerator: 'CmdOrCtrl+A',
        role: 'selectall'
    }]
}, {
    label: 'View',
    submenu: [{
        label: 'Reload',
        accelerator: 'CmdOrCtrl+R',
        click: function (item, focusedWindow) {
            if (focusedWindow) {
                // on reload, start fresh and close any old
                // open secondary windows
                if (focusedWindow.id === 1) {
                    BrowserWindow.getAllWindows().forEach(function (win) {
                        if (win.id > 1) {
                            win.close()
                        }
                    })
                }
                focusedWindow.reload()
            }
        }
    }, {
        label: 'Toggle Full Screen',
        accelerator: (function () {
            if (process.platform === 'darwin') {
                return 'Ctrl+Command+F'
            } else {
                return 'F11'
            }
        })(),
        click: function (item, focusedWindow) {
            if (focusedWindow) {
                focusedWindow.setFullScreen(!focusedWindow.isFullScreen())
            }
        }
    }, {
        label: 'Toggle Developer Tools',
        accelerator: (function () {
            if (process.platform === 'darwin') {
                return 'Alt+Command+I'
            } else {
                return 'Ctrl+Shift+I'
            }
        })(),
        click: function (item, focusedWindow) {
            if (focusedWindow) {
                focusedWindow.toggleDevTools()
            }
        }
    }, {
        type: 'separator'
    }, {
        label: 'App Menu Demo',
        click: function (item, focusedWindow) {
            if (focusedWindow) {
                const options = {
                    type: 'info',
                    title: 'Application Menu Demo',
                    buttons: ['Ok'],
                    message: 'This demo is for the Menu section, showing how to create a clickable menu item in the application menu.'
                }
                electron.dialog.showMessageBox(focusedWindow, options, function () {})
            }
        }
    }]
}, {
    label: 'Window',
    role: 'window',
    submenu: [{
        label: 'Minimize',
        accelerator: 'CmdOrCtrl+M',
        role: 'minimize'
    }, {
        label: 'Close',
        accelerator: 'CmdOrCtrl+W',
        role: 'close'
    }, {
        type: 'separator'
    }, {
        label: 'Reopen Window',
        accelerator: 'CmdOrCtrl+Shift+T',
        enabled: false,
        key: 'reopenMenuItem',
        click: function () {
            app.emit('activate')
        }
    }]
}, {
    label: 'Help',
    role: 'help',
    submenu: [{
        label: 'Learn More',
        click: function () {
            electron.shell.openExternal('http://electron.atom.io')
        }
    }]
}]


function addUpdateMenuItems(items, position) {
    if (process.mas) return

    const version = electron.app.getVersion()
    let updateItems = [{
        label: `Version ${version}`,
        enabled: false
    }, {
        label: 'Checking for Update',
        enabled: false,
        key: 'checkingForUpdate'
    }, {
        label: 'Check for Update',
        visible: false,
        key: 'checkForUpdate',
        click: function () {
            require('electron').autoUpdater.checkForUpdates()
        }
    }, {
        label: 'Restart and Install Update',
        enabled: true,
        visible: false,
        key: 'restartToUpdate',
        click: function () {
            require('electron').autoUpdater.quitAndInstall()
        }
    }]

    items.splice.apply(items, [position, 0].concat(updateItems))
}

function findReopenMenuItem() {
    const menu = Menu.getApplicationMenu()
    if (!menu) return

    let reopenMenuItem
    menu.items.forEach(function (item) {
        if (item.submenu) {
            item.submenu.items.forEach(function (item) {
                if (item.key === 'reopenMenuItem') {
                    reopenMenuItem = item
                }
            })
        }
    })
    return reopenMenuItem
}

if (process.platform === 'darwin') {
    const name = electron.app.getName()
    template.unshift({
        label: name,
        submenu: [{
            label: `About ${name}`,
            role: 'about'
        }, {
            type: 'separator'
        }, {
            label: 'Services',
            role: 'services',
            submenu: []
        }, {
            type: 'separator'
        }, {
            label: `Hide ${name}`,
            accelerator: 'Command+H',
            role: 'hide'
        }, {
            label: 'Hide Others',
            accelerator: 'Command+Alt+H',
            role: 'hideothers'
        }, {
            label: 'Show All',
            role: 'unhide'
        }, {
            type: 'separator'
        }, {
            label: 'Quit',
            accelerator: 'Command+Q',
            click: function () {
                app.quit()
            }
        }]
    })

    // Window menu.
    template[3].submenu.push({
        type: 'separator'
    }, {
        label: 'Bring All to Front',
        role: 'front'
    })

    addUpdateMenuItems(template[0].submenu, 1)
}

if (process.platform === 'win32') {
    const helpMenu = template[template.length - 1].submenu
    addUpdateMenuItems(helpMenu, 0)
}

app.on('browser-window-created', function () {
    let reopenMenuItem = findReopenMenuItem()
    if (reopenMenuItem) reopenMenuItem.enabled = false
})

// app.on('window-all-closed', function () {
//     let reopenMenuItem = findReopenMenuItem()
//     if (reopenMenuItem) reopenMenuItem.enabled = true
// })

let timeoutScheduled = Date.now();

app.on('ready', () => {
    const menu = Menu.buildFromTemplate(template)
    Menu.setApplicationMenu(menu)
    mainWindow = new BrowserWindow({
        width: 1024,
        height: 768,
        backgroundColor: '#FFFFFF',
        transparent: true,
        frame: true,
        titleBarStyle: 'hidden'
    })
    // mainWindow.loadURL('https://github.com')
    mainWindow.loadURL(url.format({
        pathname: path.join(__dirname, 'index.html'),
        protocol: 'file:',
        slashes: true
    }))
    mainWindow.on('close', function () {
        mainWindow = null
        console.log('close')
    })

    ipcMain.on('open-file-dialog', function (event) {
        dialog.showOpenDialog({
            filters: [{
                name: 'Excel',
                extensions: ['xlsx', 'xls', 'csv', 'ods']
            }],
            properties: ['multiSelections']
        }, function (files) {
            if (files) {
                //Thread
                setImmediate(() => {
                    event.sender.send('selected-directory', files)
                    event.sender.send('result', handle(files))
                })
            }
        })
    })

    ipcMain.on('click-filterBtn', function (event, symbol, times) {
        console.log('Symbol=' + symbol + ' times=' + times)
        if (times == '' || symbol == '') {} else {
            setImmediate(() => {
                event.sender.send('result', filter(result, symbol, times))
            })
        }
    })

})



let result = new Array()

function arrTo2arr(array) {
    let arr1 = new Array()
    let arr2 = new Array()
    let num = 1
    let temp = ""
    let size = array.length
    for (let i = 0; i < size; i++) {
        for (let j = i + 1; j < size; j++) {
            if (array[i] == array[j]) {
                temp = array[j]
                array[j] = array[size - 1]
                array[size - 1] = temp
                size--
                j--
                num++
            }
        }
        arr1[i] = num
        num = 1
        arr2[i] = array[i]
    }
    return {
        arr1: arr1,
        arr2: arr2
    }
}

function handle(files) {
    result = new Array()
    for (let f of files) {
        let workbook = xlsx.readFile(f);
        let sheet_name_list = workbook.SheetNames;
        sheet_name_list.forEach(function (y) { /* iterate through sheets */
            let worksheet = workbook.Sheets[y];
            for (let z in worksheet) {
                /* all keys that do not begin with "!" correspond to cell addresses */
                if (z[0] === '!') continue;
                result.push(worksheet[z].v)
                console.log(y + "!" + z + "=" + JSON.stringify(worksheet[z].v));
            }
        });
    }
    // return Array.from(new Set(result))
    return repeatArr(result)
}

//filter by t
function filter(array, symbol, t) {
    let {
        arr1,
        arr2
    } = arrTo2arr(array)
    let s = ''
    for (let i = 0; i < arr2.length; i++) {
        switch (symbol) {
            case '1': //>
                if (arr1[i] > t) {
                    s += arr2[i] + ','
                }
                break;
            case '2': //<
                if (arr1[i] < t) {
                    s += arr2[i] + ','
                }
                break;
            default: //=
                if (arr1[i] == t) {
                    s += arr2[i] + ','
                }
        }
    }
    s += '</br>'
    return s
}

function repeatArr(array) {
    let {
        arr1,
        arr2
    } = arrTo2arr(array)
    let s = ''
    let max = arr2.reduce((x, y) => (x > y) ? x : y)
    let ts = Array.from(new Set(arr1))
    ts.sort().reverse()
    for (let t of ts) {
        s += t + ':'
        for (let i = 0; i < arr2.length; i++) {
            if (arr1[i] == t) {
                s += arr2[i] + ','
            }
            console.log("----------------------element:" + arr1[i] + " repeat number:" + arr2[i] + ',')
        }
        s += '</br>'
    }
    return s
}