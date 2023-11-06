import express from 'express';
import path from 'path';
import fs from 'fs';
import cors from 'cors';
import bodyParser from 'body-parser';
import { TemplateHandler } from 'easy-template-x';
import { __dirname } from "./global.js";
console.log(__dirname);
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
    const fileName = `Hasil-${new Date().toISOString().slice(0, 10)}-${data.nama.replace(" ", "_")}.docx`
    const filePath = path.join(__dirname, `/public/doc/hasil/${fileName}`);
    fs.writeFileSync(filePath, doc);

    res.download(`${filePath}`, fileName, (err) => {
        if (err) {
            console.error(err);
            res.status(500).send('Internal server error');
        }
        fs.unlinkSync(`${filePath}`);
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
    const fileName = `Invoice-${new Date().toISOString().slice(0, 10)}-${data.nama.replace(" ", "_")}.docx`
    const filePath = path.join(__dirname, `/public/doc/invoice/${fileName}`);
    fs.writeFileSync(filePath, doc);

    res.download(`${filePath}`, fileName, (err) => {
        if (err) {
            console.error({ err });
            res.status(500).send('Internal server error');
        }
        fs.unlinkSync(`${filePath}`);
    });
});


app.listen(port, () => console.log(`Example app listening on port ${port}!`))