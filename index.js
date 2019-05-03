const xlsx = require('node-xlsx')
const fs = require('fs')

function readFile(file) {
  return xlsx.parse(file)[0].data
}

async function writeFile(file, name, data) {
  return new Promise((resolve, reject) => {
    const buffer = xlsx.build([{ name, data }])
    fs.writeFile(file, buffer, (err) => {
      if (err) {
        throw err
      }
      resolve(true)
    })
  })
}

async function start() {
  const test1 = readFile('./test1.xlsx')
  const test2 = readFile('./test2.xlsx')

  let itemNameArr = []
  for (let i=1; i<test2.length; i++) {
    itemNameArr.push(test2[i][0])
  }

  let data = [
    [
      '项目名',
      'test1',
      'test2',
      '备注'
    ]
  ]

  for (let i=1; i<test1.length; i++) {
    let itemName = test1[i][0]
    let index = itemNameArr.indexOf(itemName) + 1
    let testData = test2[index]
    for (let j=1; j<test1[i].length; j++) {
      if (test1[i][j] !== testData[j]) {
        data.push([
          itemName,
          test1[i][j],
          testData[j],
          test1[0][j]
        ])
      }
    }
  }

  console.log(data)

  const result = await writeFile('./result.xlsx', 'result', data)
}

start()
