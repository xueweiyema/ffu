'use strict'
const electron = require('electron')
const xlsx = require('xlsx')
const os = require('os')
const url = require('url')
const path = require('path')
const fs = require('fs')
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
        }
        // , {
        //     label: 'App Menu Demo',
        //     click: function (item, focusedWindow) {
        //         if (focusedWindow) {
        //             const options = {
        //                 type: 'info',
        //                 title: 'Application Menu Demo',
        //                 buttons: ['Ok'],
        //                 message: 'This demo is for the Menu section, showing how to create a clickable menu item in the application menu.'
        //             }
        //             electron.dialog.showMessageBox(focusedWindow, options, function () {})
        //         }
        //     }
        // }
    ]
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
            label: 'About',
            click: function (item, focusedWindow) {
                if (focusedWindow) {
                    const options = {
                        type: 'info',
                        title: 'filter for unique from EXCEL',
                        buttons: ['Ok'],
                        message: 'programmed by 野马VC'
                    }
                    electron.dialog.showMessageBox(focusedWindow, options, function () {})
                }
            }
        }
        // , {
        //     label: 'Learn More',
        //     click: function () {
        //         electron.shell.openExternal('http://electron.atom.io')
        //     }
        // }
    ]
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

let timeoutScheduled = Date.now()

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
            } else {
                event.sender.send('dismissed', 'dismissed')
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

    ipcMain.on('save-file-dialog', function (event, r) {
        dialog.showSaveDialog({
            title: 'ffuOut',
            filters: [{
                name: 'Normal text file',
                extensions: ['txt']
            }]
        }, function (file) {
            if (file) {
                //Thread
                setImmediate(() => {
                    fs.writeFile(file, r.replace(/<\/br>/g, "\r\n"), (err) => {
                        if (err) throw err
                        console.log('It\'s saved!')
                    })
                    event.sender.send('dismissed', 'dismissed')
                })
            } else {
                event.sender.send('dismissed', 'dismissed')
            }
        })
    })

    ipcMain.on('save-excel-dialog', function (event, r) {
        dialog.showSaveDialog({
            title: 'ffuOut',
            filters: [{
                name: 'Excel',
                extensions: ['xlsx']
            }]
        }, function (file) {
            if (file) {
                process.on('uncaughtException', function (err) {
                    event.sender.send('error', err.message)
                })
                //Thread
                setImmediate(() => {
                    // 导出 Excel
                    try {
                        process.nextTick(function () {
                            SaveExcel(r, file)
                        })
                    } catch (e) {
                        console.error('Error caught by catch block:', e);
                    }
                    event.sender.send('dismissed', 'dismissed')
                })
            } else {
                event.sender.send('dismissed', 'dismissed')
            }
        })
    })

})

function SaveExcel(r, path) {
    let arr1 = r.split("</br>")

    let objectArray = new Array()

    for (let a of arr1) {
        let v = a.split(':')[1]
        console.log('v===' + v)
        if (v == "undefined" || v == "" || v == null) {} else {
            let arr2 = v.split(',')
            for (let name of arr2) {
                let object = new Object()
                object.times = a.split(':')[0]
                object.name = name
                objectArray.push(object)
            }
        }
    }
    let _data = JSON.parse(JSON.stringify(objectArray));

    var _headers = ['times', 'name']
    // var _data = [{
    //         times: '1',
    //         name: 'test1'
    //     },
    //     {
    //         times: '2',
    //         name: 'a'
    //     },
    //     {
    //         times: '3',
    //         name: 'c'
    //     }
    // ]
    var headers = _headers
        // 为 _headers 添加对应的单元格位置
        // [ { v: 'id', position: 'A1' },
        //   { v: 'name', position: 'B1' },
        //   { v: 'age', position: 'C1' },
        //   { v: 'country', position: 'D1' },
        //   { v: 'remark', position: 'E1' } ]
        .map((v, i) => Object.assign({}, {
            v: v,
            position: String.fromCharCode(65 + i) + 1
        }))
        // 转换成 worksheet 需要的结构
        // { A1: { v: 'id' },
        //   B1: { v: 'name' },
        //   C1: { v: 'age' },
        //   D1: { v: 'country' },
        //   E1: { v: 'remark' } }
        .reduce((prev, next) => Object.assign({}, prev, {
            [next.position]: {
                v: next.v
            }
        }), {})
    var data = _data
        // 匹配 headers 的位置，生成对应的单元格数据
        // [ [ { v: '1', position: 'A2' },
        //     { v: 'test1', position: 'B2' },
        //     { v: '30', position: 'C2' },
        //     { v: 'China', position: 'D2' },
        //     { v: 'hello', position: 'E2' } ],
        //   [ { v: '2', position: 'A3' },
        //     { v: 'test2', position: 'B3' },
        //     { v: '20', position: 'C3' },
        //     { v: 'America', position: 'D3' },
        //     { v: 'world', position: 'E3' } ],
        //   [ { v: '3', position: 'A4' },
        //     { v: 'test3', position: 'B4' },
        //     { v: '18', position: 'C4' },
        //     { v: 'Unkonw', position: 'D4' },
        //     { v: '???', position: 'E4' } ] ]
        .map((v, i) => _headers.map((k, j) => Object.assign({}, {
            v: v[k],
            position: String.fromCharCode(65 + j) + (i + 2)
        })))
        // 对刚才的结果进行降维处理（二维数组变成一维数组）
        // [ { v: '1', position: 'A2' },
        //   { v: 'test1', position: 'B2' },
        //   { v: '30', position: 'C2' },
        //   { v: 'China', position: 'D2' },
        //   { v: 'hello', position: 'E2' },
        //   { v: '2', position: 'A3' },
        //   { v: 'test2', position: 'B3' },
        //   { v: '20', position: 'C3' },
        //   { v: 'America', position: 'D3' },
        //   { v: 'world', position: 'E3' },
        //   { v: '3', position: 'A4' },
        //   { v: 'test3', position: 'B4' },
        //   { v: '18', position: 'C4' },
        //   { v: 'Unkonw', position: 'D4' },
        //   { v: '???', position: 'E4' } ]
        .reduce((prev, next) => prev.concat(next))
        // 转换成 worksheet 需要的结构
        //   { A2: { v: '1' },
        //     B2: { v: 'test1' },
        //     C2: { v: '30' },
        //     D2: { v: 'China' },
        //     E2: { v: 'hello' },
        //     A3: { v: '2' },
        //     B3: { v: 'test2' },
        //     C3: { v: '20' },
        //     D3: { v: 'America' },
        //     E3: { v: 'world' },
        //     A4: { v: '3' },
        //     B4: { v: 'test3' },
        //     C4: { v: '18' },
        //     D4: { v: 'Unkonw' },
        //     E4: { v: '???' } }
        .reduce((prev, next) => Object.assign({}, prev, {
            [next.position]: {
                v: next.v
            }
        }), {});
    // 合并 headers 和 data
    var output = Object.assign({}, headers, data);
    // 获取所有单元格的位置
    var outputPos = Object.keys(output);
    // 计算出范围
    var ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];
    // 构建 workbook 对象
    var wb = {
        SheetNames: ['mySheet'],
        Sheets: {
            'mySheet': Object.assign({}, output, {
                '!ref': ref
            })
        }
    };
    // 导出 Excel
    xlsx.writeFile(wb, path)
}

// process.on('uncaughtException', function (err) {
//     ipcMain.send('error', err)
//     console.error('Error caught in uncaughtException event:', err);
// });


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
    s = s.substring(0, s.length - 1)
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
        s = s.substring(0, s.length - 1)
        s += '</br>'
    }
    return s
}