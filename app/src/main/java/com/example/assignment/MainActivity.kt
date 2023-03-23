package com.example.assignment

import android.content.ContentResolver
import android.content.ContentValues
import android.net.Uri
import android.os.Build
import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.os.Environment
import android.provider.MediaStore
import android.util.Log
import android.widget.Button
import android.widget.EditText
import android.widget.ImageView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.annotation.RequiresApi
import androidx.core.content.FileProvider

import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

class MainActivity : AppCompatActivity() {

    lateinit var imageUri : Uri
    lateinit var imageView : ImageView
    val tag = "checkval"

    private val contract = registerForActivityResult(ActivityResultContracts.TakePicture()){
        imageView.setImageURI(null)
        imageView.setImageURI(imageUri)
    }
    var rowNumber = 0



    @RequiresApi(Build.VERSION_CODES.Q)
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)


        imageView = findViewById(R.id.imageView)
        imageUri = createImageUri()!!

        val take_pic = findViewById<Button>(R.id.button)

        take_pic.setOnClickListener{
            contract.launch(imageUri)
        }




        val contentValues = ContentValues().apply {
            put(MediaStore.MediaColumns.DISPLAY_NAME,"MyExcelFile.xlsx")
            put(MediaStore.MediaColumns.MIME_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        }
        val resolver = contentResolver
        val uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, contentValues)


        val submit = findViewById<Button>(R.id.submit_button)
        submit.setOnClickListener{


            val buyerid : EditText = findViewById(R.id.buyer)
            val buyer = buyerid.text.toString()

            val styleNumid : EditText = findViewById(R.id.style_number)
            val StyleNumber = styleNumid.text.toString()

            val fabricid : EditText = findViewById(R.id.fabric)
            val fabric = fabricid.text.toString()

            val patternNumberid : EditText = findViewById(R.id.pattern_number)
            val patternNum = patternNumberid.text.toString()

            val values = values(buyer,StyleNumber,fabric,patternNum,imageUri)

            val file = File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), "MyExcelFile.xlsx")
            if(!file.exists()){

                //Wrap it in coroutine
                    if(uri!=null){
                        CoroutineScope(Dispatchers.Main).launch {
                            createNewSheet(uri,resolver,rowNumber,values)
                            Toast.makeText(this@MainActivity,"File is created in downloads folder",Toast.LENGTH_SHORT).show()
                            rowNumber++
                        }

                    }


            }else{
                Toast.makeText(this,"File already exist",Toast.LENGTH_SHORT).show()
//                CoroutineScope(Dispatchers.IO).launch {
//                    addIntoCurrentSheet(uri,resolver,rowNumber,values,file)
//                }
                val inputStream = contentResolver.openInputStream(uri!!)
                Log.d(tag,"Got the inputstream")
                val workbook = XSSFWorkbook(inputStream)
                Log.d(tag,"Got the workbook")
                val sheet = workbook.getSheetAt(0)

                val row = sheet.createRow(sheet.lastRowNum + 1)
                Log.d(tag,"Got the row")
                row.createCell(0).setCellValue(values.buyer)
                row.createCell(1).setCellValue(values.styleNumber)
                row.createCell(2).setCellValue(values.fabric)
                row.createCell(3).setCellValue(values.patternNumber)
                row.createCell(4).setCellValue(values.ImageUri.toString())
                Log.d(tag,"added the values")
                if (inputStream != null) {
                    inputStream.close()
                }
                Log.d(tag,"Closing the stream")
                val outputStream = contentResolver.openOutputStream(uri!!)
                workbook.write(outputStream)
                if (outputStream != null) {
                    outputStream.close()
                }


            }

            }


        }




        suspend fun createNewSheet(
        uri: Uri?,
        resolver: ContentResolver,
        rowNumber: Int,
        values: values
         ){
             if (uri != null) {
                 resolver.openOutputStream(uri).use{outputStream->

                val workbook = XSSFWorkbook()
                val sheet = workbook.createSheet("My Worksheet")
                     val row = sheet.createRow(0)




                    row.createCell(0).setCellValue("Buyer")
                    row.createCell(1).setCellValue("Style Number")
                    row.createCell(2).setCellValue("Fabric")
                    row.createCell(3).setCellValue("Pattern Number")
                    row.createCell(4).setCellValue("Image Uri")



                val dataRow = sheet.createRow(1)
                dataRow.createCell(0).setCellValue(values.buyer)
                dataRow.createCell(1).setCellValue(values.styleNumber)
                dataRow.createCell(2).setCellValue(values.fabric)
                dataRow.createCell(3).setCellValue(values.patternNumber)
                dataRow.createCell(4).setCellValue(values.ImageUri.toString())

                workbook.write(outputStream)
                if (outputStream != null) {
                    outputStream.close()
                }

            }
        }


    }

    suspend fun addIntoCurrentSheet(
        uri: Uri?,
        resolver: ContentResolver,
        rowNumber: Int,
        values: values,
        file: File
    ) {

        val inputStream = FileInputStream(file)
        val workbook = XSSFWorkbook(inputStream)
        val sheet = workbook.getSheetAt(0)

        val row = sheet.createRow(sheet.lastRowNum + 1)

        row.createCell(0).setCellValue(values.buyer)
        row.createCell(1).setCellValue(values.styleNumber)
        row.createCell(2).setCellValue(values.fabric)
        row.createCell(3).setCellValue(values.patternNumber)
        row.createCell(4).setCellValue(values.ImageUri.toString())

        inputStream.close()

        val outputStream = FileOutputStream(file)
        workbook.write(outputStream)
        outputStream.close()


    }


    private fun createImageUri(): Uri? {
        val image = File(applicationContext.filesDir,"camera_photos.png")

        return FileProvider.getUriForFile(applicationContext,"com.example.assignment.FileProvider",image)

    }

}




