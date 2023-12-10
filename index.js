import express from 'express';
import path from 'path';
import fs from 'fs';
import cors from 'cors';
import bodyParser from 'body-parser';
import { TemplateHandler } from 'easy-template-x';
import XlsxTemplate from 'xlsx-template';
import { __dirname } from "./global.js";

const app = express()
const port = 3000

app.use(express.urlencoded({ extended: true }));
app.use(cors());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

app.post('/generate-word-document', async (req, res) => {
    // 1. read template file
    const templateFile = fs.readFileSync('./public/doc/template/LaporanHasil.docx');

    // 2. process the template
    const data = req.body
    const handler = new TemplateHandler();
    const doc = await handler.process(templateFile, data);

    // 3. send output
    const fileName = `${new Date().toISOString().slice(0, 10)}-${data.nama.replace(" ", "_")}.docx`
    const filePath = path.join(__dirname, `/public/doc/hasil/${fileName}`);
    fs.writeFileSync(filePath, doc);
    console.log(filePath);
    res.download(`${filePath}`, fileName, (err) => {
        if (err) {
            console.error(err);
            res.status(500).send('Internal server error');
            fs.unlinkSync(`${filePath}`);
        }
    });
});

app.post('/generate-invoice', async (req, res) => {
    // 1. read template file
    const templateFile = fs.readFileSync('./public/doc/template/invoice.docx');

    // 2. process the template
    const data = req.body
    const handler = new TemplateHandler();
    const doc = await handler.process(templateFile, data);

    // 3. send output
    const fileName = `${new Date().toISOString().slice(0, 10)}-${data.nama.replace(" ", "_")}.docx`
    const filePath = path.join(__dirname, `/public/doc/invoice/${fileName}`);
    fs.writeFileSync(filePath, doc);
    res.download(`${filePath}`, fileName, (err) => {
        if (err) {
            console.error({ err });
            res.status(500).send('Internal server error');
            fs.unlinkSync(`${filePath}`);
        }
    });
});

app.post('/generate-excel', async (req, res) => {

    fs.readFile(path.join('./public/xlsx/template/bon.xlsx'), function (err, data) {

        // Create a template
        var template = new XlsxTemplate(data);

        // Replacements take place on first sheet
        var sheetNumber = 1;

        // Set up some placeholder values matching the placeholders in the template
        var values = {
            tanggal: '2023-11-22',
            penerima: 'gilang',
            jenis_jasa: 'cek tanah',
            total: 20000,
            tgltanda: "bandung,24 mei 2024"
        };

        // Perform substitution
        template.substitute(sheetNumber, values);

        // Get binary data
        var data = template.generate();
        const fileName = `${new Date().toISOString().slice(0, 10)}-${values.penerima.replace(" ", "_")}.xlsx`
        const filePath = path.join(__dirname, `/public/xlsx/output/${fileName}`);

        fs.writeFileSync(filePath, data, 'binary');
        res.download(`${filePath}`, fileName, (err) => {
            if (err) {
                console.error({ err });
                res.status(500).send('Internal server error');
                fs.unlinkSync(`${filePath}`);
            }
            console.log("download berhasil");
        });
    })
})

app.listen(port, () => console.log(`Example app listening on port ${port}!`))