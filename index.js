const mongoose = require("mongoose");
const xlsx = require("xlsx");

// MongoDB connection URI
const uri = "mongodb://localhost:27017/invoice-system"; // Replace with your database name

// Define the schema
const ProductSchema = new mongoose.Schema({
  medicineName: { type: String },
  batch: { type: String },
  expiryDate: { type: String },
  price: { type: Number },
  quantity: { type: Number },
  sgst: { type: Number },
  hsn_code: { type: String },
  createdAt: { type: String },
  createdBy: { type: String },
  updatedBy: { type: String },
  updatedOn: { type: String },
  supplierName: { type: String },
  supplierAddress: { type: String },
  supplierPhone: { type: String },
  mid: { type: String },
});

const Product = mongoose.model("medicine-managements", ProductSchema);

// Connect to MongoDB
mongoose
  .connect(uri, {
    useNewUrlParser: true,
    useUnifiedTopology: true,
  })
  .then(() => {
    console.log("Connected to MongoDB");
    // Proceed to read and check Excel data
    insertExcelData("/Users/garygabreil/Downloads/test.xls"); // Update with the correct path
  })
  .catch((error) => {
    console.error("Error connecting to MongoDB:", error);
  });

// Function to clean up the text (remove special characters, convert to uppercase, replace spaces with underscores)
function cleanText(text) {
  return text
    .replace(/[^a-zA-Z0-9\s]/g, "") // Removes special chars except spaces
    .trim()
    .replace(/\s+/g, "_") // Replace spaces with underscores
    .toUpperCase();
}

// Function to check if a value is numeric
function isNumeric(value) {
  return !isNaN(value) && isFinite(value);
}

// Function to extract numeric value from a string
function extractGST(gstString) {
  if (!gstString) return null; // Return null if gstString is undefined or null
  const match = gstString.match(/(\d+)/);
  return match ? parseFloat(match[0]) : null; // Return the number or null if not found
}

// Function to generate a unique mid
function generateUniqueNumber() {
  return Math.floor(100000 + Math.random() * 900000); // Generates a random 6-digit number
}

// Function to read data from the Excel file and insert it into MongoDB
function insertExcelData(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const rows = [];
    const range = xlsx.utils.decode_range(worksheet["!ref"]); // Get the range of the sheet

    for (let rowNum = range.s.r + 1; rowNum <= range.e.r; rowNum++) {
      // Start from the second row (after header)
      const product_name = worksheet[`B${rowNum}`]
        ? worksheet[`B${rowNum}`].v
        : null; // Column B
      const mrp = worksheet[`L${rowNum}`] ? worksheet[`L${rowNum}`].v : null; // Column L
      const suppliername = worksheet[`J${rowNum}`]
        ? worksheet[`J${rowNum}`].v
        : null; // Column J
      const gstString = worksheet[`R${rowNum}`]
        ? worksheet[`R${rowNum}`].v
        : null; // Column R

      // Only process rows where `mrp` is numeric
      if (product_name && isNumeric(mrp) && suppliername) {
        const cleanProduct = {
          medicineName: cleanText(product_name || ""),
          price: parseFloat(mrp), // Convert `mrp` to a number
          supplierName: cleanText(suppliername || ""),
          sgst: extractGST(gstString), // Extract GST value
          mid: generateUniqueNumber(), // Generate a unique mid for each product
        };
        rows.push(cleanProduct);
      } else {
        console.log(`Skipping row ${rowNum} due to invalid data: `, {
          product_name,
          mrp,
          suppliername,
          gstString,
        });
      }
    }

    // Insert the valid rows into MongoDB
    if (rows.length > 0) {
      Product.insertMany(rows)
        .then(() => {
          console.log("Data successfully inserted into MongoDB");
        })
        .catch((error) => {
          console.error("Error inserting data into MongoDB:", error);
        });
    } else {
      console.log("No valid data to insert.");
    }
  } catch (error) {
    console.error("Error reading Excel file:", error);
  }
}
