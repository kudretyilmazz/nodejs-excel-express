// Import Express
const express = require("express");

// Import ExcelJs
const exceljs = require("exceljs");

// Variables
const app = express();
const PORT = 3000;

// Server Go Live!
app.listen(PORT, () => {
	console.log(`running on ${PORT}`);
});

// Mock Data
const data = [
	{
		name: "Pencil",
		price: 5,
		updatedAt: "10.10.2010",
	},
	{
		name: "Book",
		price: 10,
		updatedAt: "08.12.2012",
	},
	{
		name: "Eraser",
		price: 2,
		updatedAt: "10.10.2010",
	},
];

// Yeni bir excel dökümanı oluşturuyoruz.
const workbook = new exceljs.Workbook();

// Excel dökümanı üzerinde bir worksheet oluşturuyoruz.
const worksheet = workbook.addWorksheet("Worksheet");

// Excel Columns
worksheet.columns = [
	{
		header: "Name",
		key: "name",
		width: 20,
	},
	{
		header: "Price",
		key: "price",
		width: 20,
	},
	{
		header: "Last Updated Date",
		key: "updatedAt",
		width: 20,
	},
];

// Add Data
worksheet.addRows(data);

app.get("/export", (req, res) => {
	// Set Header
	res.setHeader(
		"Content-Type",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
	);
	res.setHeader("Content-Disposition", "attachment; filename=data.xlsx");

	return workbook.xlsx.write(res).then(() => {
		res.status(200).end();
	});
});
