const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const { Parser } = require('json2csv');
const fs = require('fs');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));

app.post('/convert', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    let jsonData = XLSX.utils.sheet_to_json(worksheet);

    // Modify each row to add the desired values to specific columns
    jsonData = jsonData.map(row => {
        const title = row['Title'];
        let option1Value = row['Option1 Value'];
        option1Value = option1Value.replace(/tester|Tester/g, 'TSTR');
        const sku = row['Variant SKU'];
        const barcode = row['Variant Barcode'];
        const concatenatedTitle = title + ' - ' + option1Value;
        const brand = title.split('-')[0].trim();

        const matches = concatenatedTitle.match(/-([^]+?)-/); 
        const extractedWord = matches ? matches[1].trim() : '';

        const size = parseInt(row['Option1 Value'].replace('/ml|Ml/g',''));
        const volume = (size * 0.033814).toFixed(2);
        const imageUrl = row['Variant Image'];
        const productPrice = Math.ceil(row['Variant Price'] * 1.10 + 15);

        const testerDescription = `<div style=""><font color="#111820" face="Arial, sans-serif">Testers often come with brown or no box and no cap.</font></div><font rwr="1" size="4" style="font-family:Arial"><div style=""><font color="#111820" face="Arial, sans-serif">1. US seller fast shipping</font></div><div style=""><font color="#111820" face="Arial, sans-serif">2. If you have any questions do not hesitate to ask</font></div><div style=""><font color="#111820" face="Arial, sans-serif">3. All items are 100% authentic olfactory guarantee</font></div><div style=""><font color="#111820" face="Arial, sans-serif">4. Shipped by UPS, FedEx&nbsp;or USPS . No international shipping</font></div></font>`;
        const normalDescription = `<font rwr="1" size="4" style="font-family:Arial"><div style=""><font color="#111820" face="Arial, sans-serif">1. US seller fast shipping</font></div><div style=""><font color="#111820" face="Arial, sans-serif">2. If you have any questions do not hesitate to ask</font></div><div style=""><font color="#111820" face="Arial, sans-serif">3. All items are 100% authentic olfactory guarantee</font></div><div style=""><font color="#111820" face="Arial, sans-serif">4. Shipped by UPS, FedEx&nbsp;or USPS . No international shipping</font></div></font>`
        
        return {
            ...row,
            '*Action(SiteID=US|Country=US|Currency=USD|Version=1193|CC=UTF-8)': 'Add',
            'CustomLabel': sku,
            '*Category': 112661,
            '*Title': concatenatedTitle,
            '*ConditionID': 1000,
            '*C:Brand': brand,
            '*C:Fragrance Name': extractedWord,
            '*C:Type': 'Perfume',
            '*C:Volume': volume + " fl oz",
            'PicURL': imageUrl,
            '*Description': (concatenatedTitle.includes('TSTR'))? testerDescription : normalDescription,
            '*Format': 'FixedPrice',
            '*Duration': 'GTC',
            '*StartPrice': productPrice,
            '*Quantity' : 1,
            'ImmediatePayRequired': 1,
            '*Location' : 'IN, USA',
            'ShippingType' : 'Flat',
            'ShippingService-1:FreeShipping' : 1,
            'ShippingService-1:Option' : 'UPSGround',
            '*DispatchTimeMax': 2,
            '*ReturnsAcceptedOption': 'ReturnsNotAccepted',
            'P:UPC' : barcode,
        };
    });
    
    const fields = [
        '*Action(SiteID=US|Country=US|Currency=USD|Version=1193|CC=UTF-8)',
        'CustomLabel',
        '*Category',
        'StoreCategory',
        '*Title',
        'Subtitle',
        'Relationship',
        'RelationshipDetails',
        '*ConditionID',
        '*C:Brand',
        '*C:Fragrance Name',
        '*C:Type',
        '*C:Volume',
        '*C:Scent',
        '*C:Formulation',
        'C:Size Type',
        'C:Country/Region of Manufacture',
        '*C:MPN',
        'C:Features',
        'C:Bundle Description',
        'C:California Prop 65 Warning',
        'C:Custom Bundle',
        'C:Expiration Date',
        'PicURL',
        'GalleryType',
        '*Description',
        '*Format',
        '*Duration',
        '*StartPrice',
        'BuyItNowPrice',
        '*Quantity',
        'PayPalAccepted',
        'PayPalEmailAddress',
        'ImmediatePayRequired',
        'PaymentInstructions',
        '*Location',
        'ShippingType',
        'ShippingService-1:FreeShipping',
        'ShippingService-1:Option',
        'ShippingService-1:Cost',
        'ShippingService-2:Option',
        'ShippingService-2:Cost',
        '*DispatchTimeMax',
        'PromotionalShippingDiscount',
        'ShippingDiscountProfileID',
        '*ReturnsAcceptedOption',
        'ReturnsWithinOption',
        'RefundOption',
        'ShippingCostPaidByOption',
        'AdditionalDetails',
        'P:UPC'
    ];

    const parser = new Parser({ fields });
    const csvData = parser.parse(jsonData);

    fs.unlinkSync(req.file.path); // Delete the uploaded file

    res.attachment('output.csv');
    res.send(csvData);
});


// Serve index.html for any other routes
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
