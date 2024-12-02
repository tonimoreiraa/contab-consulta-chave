import fs from 'fs/promises';
import path from 'path';
import xlsx from 'xlsx';
import puppeteer from 'puppeteer';
import { config } from 'dotenv'

config()

const INPUT_DIR = 'input';

const login = process.env.LOGIN;
const password = process.env.SENHA;

async function getAccessKeys()
{
    const accessKeys = []
    try {
        const files = await fs.readdir(INPUT_DIR);
        
        for (const filename of files) {
            
            if (!filename.endsWith('.xlsx')) {
                continue;
            }

            const filePath = path.join(INPUT_DIR, filename);
            
            try {
                const workbook = xlsx.readFile(filePath);
                
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];

                const data = xlsx.utils.sheet_to_json(worksheet);

                data.forEach(row => {
                    const value = row._2;
                    
                    if (String(value).length === 44) {
                        accessKeys.push(value)
                    }
                });

            } catch (error) {
                console.error(`Error processing file ${filename}:`, error);
            }
        }
    } catch (err) {
        console.error('Error reading directory:', err);
    }

    return accessKeys;
}

async function processExcelFiles() {
    const accessKeys = await getAccessKeys()
    const driver = await puppeteer.launch({ headless: false })
    const page = await driver.newPage()
    await page.goto('https://nfce.sefaz.al.gov.br')
    await page.waitForSelector('#sca_login')
    await page.type('#sca_login', login)
    await page.type('#sca_senha', password)
    await page.click('input[name="btn_entrar"]')
    await page.waitForSelector('#DHTMLSuite_menuItem2')

    const values = []
    for (const accessKey of accessKeys) {
        await page.goto('https://nfce.sefaz.al.gov.br/paginaConsultaDanfeDetalhado.htm')

        await page.waitForSelector('#chaveAcesso')
        await page.evaluate((accessKey) => {
            document.getElementById('chaveAcesso').value = accessKey
        }, accessKey)

        await page.click('input[type=submit]')

        await page.waitForSelector('.GeralXslt')
        const element = await page.waitForSelector('#NFe > fieldset:nth-child(4) > legend')
        const value = await element.evaluate(el => el.textContent)
        await new Promise(r => setTimeout(r, 10000))

        let parsedValue = value.trim()
        parsedValue = parsedValue.slice(parsedValue.indexOf(':') + 1, parsedValue.indexOf('(')).trim()
        values.push([accessKey, parsedValue])

        await fs.writeFile('./output.csv', values.map(v => v.join(';')).join('\n'))
    }
    await driver.close()
}

processExcelFiles()