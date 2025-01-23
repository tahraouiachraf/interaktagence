const axios = require("axios");
const XLSX = require("xlsx");
const fs = require("fs");

const apiUrl = 'https://apiovet.animalmarket.ma/api/product/list?page=1&limit=10';
const apiKey = '9db9106111f8092c4edf02d83a11f716';
const excelFilePath = './allitems.xlsx';

const fetchDataFromAPI = async () => {
    try {
        const response = await axios.get(apiUrl, {
            headers: { 'x-api-key': apiKey }
        });
        return response.data.products;
    } catch (error) {
        console.error("Erreur lors de la récupération des données de l'API :", error);
        return [];
    }
};

const updateXlsxFile = async () => {
    const products = await fetchDataFromAPI();

    if (!products || products.length === 0) {
        console.log("Aucun produit trouvé.");
        return;
    }

    if (!fs.existsSync(excelFilePath)) {
        console.error("Le fichier XLSX n'existe pas.");
        return;
    }

    const workbook = XLSX.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    rows.forEach((row, index) => {
        if (index === 0) return;

        const slugCell = row[4];
        if (slugCell) {
            const product = products.find((p) => p.slug === slugCell);
            if (product) {
                row[6] = product.link;
                console.log(`Mise à jour du produit ${product.name} (slug: ${slugCell}) avec le lien: ${product.link}`);
            }
        }
    });

    const updatedSheet = XLSX.utils.aoa_to_sheet(rows);

    workbook.Sheets[sheetName] = updatedSheet;

    const outputFilePath = './allitemsUpload.xlsx';
    XLSX.writeFile(workbook, outputFilePath);

    console.log(`Fichier mis à jour et sauvegardé sous : ${outputFilePath}`);
};

updateXlsxFile();
