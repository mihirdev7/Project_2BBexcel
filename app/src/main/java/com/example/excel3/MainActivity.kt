package com.example.excel3

import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.provider.Settings
import androidx.appcompat.app.AppCompatActivity
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import android.widget.Button
import android.widget.Toast

class MainActivity : AppCompatActivity() {
    private lateinit var recyclerView: RecyclerView
    private lateinit var exportButton: Button
    private val customers = mutableListOf<Customer>()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        recyclerView = findViewById(R.id.recyclerView)
        exportButton = findViewById(R.id.exportButton)

        customers.addAll(
            listOf(
                Customer(1, "Ronaldo", "Siuuu@gmail.com", "1234567890", "Delhi"),
                Customer(2, "Smith", "Smith@yahoo.com", "0987654321", "Mumbai"),
                Customer(3, "Michael", "michael@google.com", "1122334455", "Gujarat"),
                Customer(4, "Emily", "emily@gmail.com", "6677889900", "Korea"),
                Customer(5, "Brown", "brown@google.com", "5544332211", "Japan")
            )
        )

        recyclerView.layoutManager = LinearLayoutManager(this)
        recyclerView.adapter = CustomerAdapter(customers)

         exportButton.setOnClickListener {
            exportDataToExcel()
        }
    }

    private fun exportDataToExcel() {
        if (!checkStoragePermission()) {
            Toast.makeText(this, "Storage permission required", Toast.LENGTH_SHORT).show()
            return
        }

        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet("Customers")

         val headerRow = sheet.createRow(0)
        headerRow.createCell(0).setCellValue("ID")
        headerRow.createCell(1).setCellValue("Name")
        headerRow.createCell(2).setCellValue("Email")
        headerRow.createCell(3).setCellValue("Phone")
        headerRow.createCell(4).setCellValue("Address")

         customers.forEachIndexed { index, customer ->
            val row = sheet.createRow(index + 1)
            row.createCell(0).setCellValue(customer.id.toDouble())
            row.createCell(1).setCellValue(customer.name)
            row.createCell(2).setCellValue(customer.email)
            row.createCell(3).setCellValue(customer.phone)
            row.createCell(4).setCellValue(customer.address)
        }

         val downloadsDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)
        val filePath = File(downloadsDir, "Customers.xlsx")
        FileOutputStream(filePath).use { fileOutputStream ->
            workbook.write(fileOutputStream)
            workbook.close()
        }
        Toast.makeText(this, "Data Saved as Customers", Toast.LENGTH_LONG).show()
    }

    private fun checkStoragePermission(): Boolean {
        return if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
            if (Environment.isExternalStorageManager()) {
                true
            } else {
                val intent = Intent(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION).apply {
                    data = Uri.parse("package:$packageName")
                }
                startActivityForResult(intent, 1001)
                false
            }
        } else {
            if (checkSelfPermission(android.Manifest.permission.WRITE_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED) {
                true
            } else {
                requestPermissions(arrayOf(android.Manifest.permission.WRITE_EXTERNAL_STORAGE), 1001)
                false
            }
        }
    }
}
