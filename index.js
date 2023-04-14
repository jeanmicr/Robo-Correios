const puppeteer = require('puppeteer');
const reader = require('xlsx')

const file = reader.readFile('./Dados Welcome 2023.xlsx')
let data = []
const sheets = file.SheetNames
const sheetsNum = sheets.length
const lastSheet = sheets[sheetsNum - 1]
const dados = reader.utils.sheet_to_json(file.Sheets[lastSheet])

function lerDadosPlanilha() {
    dados.forEach(content => {
        data.push(content)
    })
    return (data)
}

function dadosJson() {
    var values = lerDadosPlanilha()
    return (values)
}
//dadosJson()

console.log('Robo de geração de etiquetas dos correios!')

async function robo() {
    const result = dadosJson()
    const browser = await puppeteer.launch({ headless: false, executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe', args: [ "--kiosk-printing" ] })
    const page = await browser.newPage()
    console.log('Acessando página dos correios')
    url = 'https://www2.correios.com.br/enderecador/encomendas/default.cfm'
    await page.goto(url)
    console.log('Preenchendo dados')
    if (result != "") {
        for (let i = 1, e = 0; i <= 1, e <= 0; i++, e++) {
            const value = result[e]
            var data = value
            var nomeDest = data['Nome']
            var ruaDest = data['RUA']
            var complementoDest = data['COMPLEMENTO']
            var cepDest = data['CEP']

            var cep = '[name="cep_' + i + '"]'
            var nome = '[name="nome_' + i + '"]'
            var numero = '[name="numero_' + i + '"]'
            var complemento = '[name="complemento_' + i + '"]'
            var uf = '[name="selUf_' + i + '"]'
            var cepDes = '[name="desCep_' + i + '"]'
            var nomeDes = '[name="desNome_' + i + '"]'
            var ruaDes = '[name="desEndereco_' + i + '"]'
            var numeroDes = '[name="desNumero_' + i + '"]'
            var complementoDes = '[name="desComplemento_' + i + '"]'
            var gerarEtiqueta = '#btGerarEtiquetas'
            var gerarAr = '#btGerarAR'

            await page.type(cep, '20031-040')
            await page.click(nome)
            await page.waitForSelector(nome)
            await page.type(nome, 'Gestão de Ativos Stone')
            await page.type(numero, '65')
            await page.type(complemento, 'Torre 2, Sala 201')
            await page.type(uf, 'RJ')

            await page.type(cepDes, cepDest)
            await page.click(nomeDes)
            await page.waitForSelector(nomeDes)
            await page.type(nomeDes, nomeDest)
            await page.click(ruaDes, { clickCount: 3 })
            await page.type(ruaDes, ruaDest)
            await page.type(numeroDes, '-')
            await page.type(complementoDes, complementoDest)
        }
        for (i = 1; i <= 4; i++) {
            var mp = '[name="mp_' + i + '"]'
            //var obsClick = '[name="desDC_' + i + '"]'
            await page.click(mp)
            //     await page.click(obsClick)
            //     await page.keyboard.press('Tab')
            //     await page.keyboard.press('Enter')
            // }

        }
        await page.click(gerarEtiqueta)

        const pageTarget = page.target()
        const newTarget = await browser.waitForTarget(target => target.opener() === pageTarget)
        const newPage = await newTarget.page()
        await newPage.waitForSelector("body")
        await newPage.close()
        //await page.evaluate(() => {window.print()})


    }
}
robo()
