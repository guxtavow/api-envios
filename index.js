const express = require('express')
const xlsx = require('xlsx')
const multer = require('multer')
const path = require('path')
const app = express()
const port = 3000

app.use(express.json())

let etiquetas = [] //aqui será a variavel que irá armazenar as etiquetas


const carregarPlanilha = () => { //função para carregar planilha
    try{
        const filePath = path.join(__dirname, './lista_etiquetas.xlsx')
        const workbook = xlsx.readFile(filePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]]
        const data = xlsx.utils.sheet_to_json(worksheet)
        etiquetas = data
        console.log('Planilha processada com sucesso!')
    }catch(error){
        console.log(error)
    }
}


const upload = multer()


app.post('/upload', upload.single('planilha'), (req, res) => {
    try{
        const workbook = xlsx.readFile(req.file.path)
        const worksheet = workbook.Sheets[workbook.SheetNames[0]]
        const data = xlsx.utils.sheet_to_json(worksheet)
        etiquetas = data
        console.log('Planilha processada com sucesso!')
        res.status(200).json({message:'Planilha processada com sucesso!'})
    }catch(err){
        console.log(err)
        res.status(200).json({message:`Planilha não processada.`, err:err.message })
    }
})


app.get('/', (req, res) => { // get de todas as etiquetas
    res.status(200).json(etiquetas)
})

app.listen(port, () => {
    console.log(`Rodando em: http://localhost:${port}`)
    carregarPlanilha()
})