const puppeteer = require('puppeteer')
const XLSX = require('xlsx')
const fs = require('fs')

async function parseLandings(filePath) {
    //Create excel worbook, sheet and pages' array
    const workbook = XLSX.readFile(filePath)
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    const pages = []
    for (let i = 2; i <= sheet['!ref'].split(':')[1].replace(/[A-Z]/g, ''); i++) {
        // check if sheet[A${i}] is defined
        if (sheet[`A${i}`]) {
            // push the value of sheet[A${i}].v to the pages array
            pages.push(sheet[`A${i}`].v)
        }
    }

    const urls300 = []
    const browser = await puppeteer.launch()
    for (const url of pages) {

        try {
            const page = await browser.newPage()
            await page.goto(url, {
                waitUntil: 'domcontentloaded',
                timeout: 60000
            })
            setTimeout(()=> {}, 2000)
            await page.waitForSelector('.pagination__item')
            const element = await page.evaluate(() => {
                return document.querySelectorAll('.pagination__item')[5].innerText
            })
            if (element === '300') {
                urls300.push({ url: url })
            } else {
                console.log(`${url}: ${element}`)
            }
        } catch (e) {
            if (e.name === 'TimeoutError') {
                console.log(`Navigation timeout of 60 seconds exceeded for ${url}`)
            }
            console.log(`${url}: < 5`)
        }
    }

    await browser.close()

    // check if the urls300.xlsx file exists
    if (fs.existsSync('urls300.xlsx')) {
        // read the file
        const workbook = XLSX.readFile('urls300.xlsx')
        // get the sheet containing the data
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        // update the sheet with the new data from urls300 array
        //This method call will add the data from the urls300 array to the sheet sheet, starting from the last row and skipping the first row of data.
        XLSX.utils.sheet_add_json(sheet, urls300, { origin: -1, skipHeader: true })
        // write the updated sheet back to the urls300.xlsx file
        XLSX.writeFile(workbook, 'urls300.xlsx')
    } else {
        // create a new workbook and sheet if the file does not exist
        const newWorkbook = XLSX.utils.book_new()
        const newSheet = XLSX.utils.json_to_sheet(urls300)
        XLSX.utils.book_append_sheet(newWorkbook, newSheet)
        XLSX.writeFile(newWorkbook, 'urls300.xlsx')
    }
}

// Path to the file, which's been reading by the program
const filePath = "C:\\Users\\User\\Desktop\\landings.xlsx"

parseLandings(filePath)









