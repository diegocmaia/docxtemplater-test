const JSZip = require('jszip')
const Docxtemplater = require('docxtemplater')

const fs = require('fs')
const path = require('path')

run = function(builderScript) {
    //Load the docx file as a binary
    const content = fs.readFileSync(
        path.resolve(__dirname, '../templates/myTemplate.docx'),
        'binary'
    )

    const zip = new JSZip(content)

    const doc = new Docxtemplater()
    doc.loadZip(zip)

    //set the templateVariables

    const data = {
        org_name: 'MyOrganization',
        transactions: [
            {
                description: 'Transaction 01',
                value: 10,
                hasDiscount: true
            },
            {
                description: 'Transaction 02',
                value: 105,
                hasDiscount: false
            },
            {
                description: 'Transaction 03',
                value: 10,
                hasDiscount: true
            },
            {
                description: 'Transaction 04',
                value: 1000,
                hasDiscount: false
            }
        ]
    }

    doc.setData(data)

    try {
        // render the document
        doc.render()
    } catch (error) {
        var e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties
        }
        console.log(JSON.stringify({ error: e }))
        // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
        throw error
    }

    const buf = doc.getZip().generate({ type: 'nodebuffer' })

    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.resolve(__dirname, '../reports/output.docx'), buf)
}

run()
