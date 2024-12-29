const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// Directory containing JSON files
const dataDir = "data";

// Get a list of all JSON files in the directory
const files = fs.readdirSync(dataDir).filter(file => file.endsWith(".json"));

let allFormattedData = [];

// Iterate through each file
files.forEach(file => {
    const filePath = path.join(dataDir, file);

    // Load and parse the JSON data
    const rawData = fs.readFileSync(filePath);
    const data = JSON.parse(rawData);

    // Format the data and add a new column with the file name
    const formattedData = data.map(item => ({
        SKU: item.skuId,
        "Product ID": item.productId,
        Name: item.name,
        Price: item.price,
        Images: (item.medias || []).map(media => media.url).join(", "), // Concatenate all image URLs
        "Short Description": item.shortDescription.replace(/<\/?[^>]+(>|$)/g, ""), // Remove HTML tags
        Description: item.detailedDescription.replace(/<\/?[^>]+(>|$)/g, ""), // Remove HTML tags
        "Source File": file // Add the source file name
    }));

    // Combine the data
    allFormattedData = allFormattedData.concat(formattedData);
});

// Create a worksheet from the combined data
const worksheet = XLSX.utils.json_to_sheet(allFormattedData);

// Create a new workbook and append the worksheet
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, "All Products");

// Write the workbook to an Excel file
XLSX.writeFile(workbook, "all_products.xlsx");

console.log("Excel file 'all_products.xlsx' has been created successfully!");
