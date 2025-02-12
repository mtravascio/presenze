import 'package:flutter/material.dart';
import 'package:get/get.dart';
import 'package:file_picker/file_picker.dart';
//import 'package:intl/date_time_patterns.dart';
import 'package:path/path.dart' as path;
import 'package:path_provider/path_provider.dart';
import 'package:excel/excel.dart';
import 'package:intl/intl.dart';
import 'dart:io';
import 'package:desktop_window/desktop_window.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  DesktopWindow.setWindowSize(Size(400, 500));
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return GetMaterialApp(
      title: 'Elaborazione File Excel Presenze',
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: HomeScreen(),
    );
  }
}

class HomeScreen extends StatelessWidget {
  final FileController fileController = Get.put(FileController());

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Elabora Presenze'),
        backgroundColor: Colors.deepOrangeAccent,
      ),
      backgroundColor: Colors.deepOrange.shade100,
      body: Center(
        child: ElevatedButton(
          onPressed: () {
            fileController.pickFile();
          },
          child: Text('Seleziona File Excel'),
        ),
      ),
    );
  }
}

class FileController extends GetxController {
  void pickFile() async {
    // Apri il selettore di file per scegliere un file Excel
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (result != null && result.files.single.path != null) {
      File file = File(result.files.single.path!);
      await processFile(file);
    }
  }

  Future<void> processFile(File file) async {
    // Leggi il file Excel usando il package 'excel'
    var bytes = await file.readAsBytes();
    var excel = Excel.decodeBytes(bytes);

    // Crea un mappa per memorizzare i dati che sono unici nella colonna nomi
    //Map<String, Set<String>> userDates = {};
    Map<String, List<Map<String, String>>> userInfo = {};

    // Supponiamo che la data sia nella colonna A e il nome nella colonna B
    for (var table in excel.tables.keys) {
      var sheet = excel.tables[table];
      if (sheet != null) {
        for (var row in sheet.rows) {
          var dataString = row[0]?.value.toString(); // Colonna Data
          var modeString = row[1]?.value.toString(); // Colonna Modalità
          var oraString = row[2]?.value.toString(); // Colonna Ora
          var distrettoString = row[3]?.value.toString(); // Colonna Distretto
          var name = row[4]?.value.toString(); // Colonna Nome e Cognome
          var ruolo = row[5]?.value.toString(); // Colonna Ruolo
          var email = row[6]?.value.toString(); // Colonna Ruolo

          if (name != null && dataString != null && dataString.isDateTime) {
            String fullName = name;

            // Formattazione della data
            DateTime data = DateTime.parse(
                dataString); // Assicurati che dataString sia nel formato corretto
            String formattedDate = DateFormat('dd/MM/yy').format(data);
            String dateString = formattedDate;

            // Aggiungi la data alla lista per quel nome, eliminando i duplicati
            if (!userInfo.containsKey(fullName)) {
              userInfo[fullName] = [];
            }
            userInfo[fullName]!.add({
              'date': dateString,
              'mode': modeString ?? '',
              'ora': oraString ?? '',
              'distretto': distrettoString ?? '',
              'ruolo': ruolo ?? '',
              'email': email ?? ''
            });
          }
        }
      }
    }

    // Ordinamento delle date per ogni nome
    userInfo.forEach((name, infoList) {
      // Ordina la lista in base alla data (convertendo la data in DateTime per il confronto)
      infoList.sort((a, b) {
        DateTime dateA = DateFormat('dd/MM/yy').parse(a['date']!);
        DateTime dateB = DateFormat('dd/MM/yy').parse(b['date']!);
        return dateA.compareTo(dateB);
      });
    });

    // Crea un nuovo file Excel di output
    final outputFile = await createOutputFile(userInfo);
    Get.snackbar('Successo', 'File Presenze.xlsx salvato in Documenti');
  }

// Ordinamento delle date per ogni nome

  Future<File> createOutputFile(
      Map<String, List<Map<String, String>>> userInfo) async {
    var excel = Excel.createExcel(); // Crea un nuovo file Excel
    var sheet = excel['Sheet1']; // Aggiungi un foglio di lavoro
//Stampa tutte le occorrenze!
/*    int row = 1;
    userInfo.forEach((name, infoList) {
      sheet.cell(CellIndex.indexByString("A$row")).value =
          TextCellValue(name); // Nome e Cognome

      // Creiamo una stringa che unisce tutte le date e modalità per ogni nome
      String datesAndModes = infoList.map((info) {
        return '${info['date']} (${info['mode']})'; // Unisce data e modalità
      }).join(
          ", "); // Unisce tutte le date/modalità con una virgola separatrice

      sheet.cell(CellIndex.indexByString("B$row")).value =
          TextCellValue(datesAndModes); // Date di collegamento + modalità
      row++;
    });
*/
//Stampa solo le occorrenze diverse!
    int row = 1;
    userInfo.forEach((name, infoList) {
      sheet.cell(CellIndex.indexByString("A$row")).value =
          TextCellValue(name); // Nome e Cognome

      // Usa un Set per eliminare i duplicati (combinazione unica di data e modalità)
      Set<String> uniqueDatesAndModes = {};

      // Aggiungiamo le combinazioni di data e modalità al Set e scrive l'ultima modalità usata il distretto di collegamento
      for (var info in infoList) {
        sheet.cell(CellIndex.indexByString("B$row")).value =
            TextCellValue(info['email'] ?? '?');
        sheet.cell(CellIndex.indexByString("C$row")).value =
            TextCellValue(info['distretto'] ?? '?');
        sheet.cell(CellIndex.indexByString("D$row")).value =
            TextCellValue(info['mode'] ?? '?');
//        String combined ='${info['date']} ${info['ora']} (${info['mode']})'; // Combina data, ora e modalità
//        String combined ='${info['date']} (${info['mode']})'; // Combina data e modalità
        String combined = '${info['date']}'; // Solo la data
        uniqueDatesAndModes
            .add(combined); // Il Set garantisce che non ci siano duplicati
      }

      // Combina tutte le voci uniche in una singola stringa separata da virgole
      String datesAndModes = uniqueDatesAndModes.join(", ");

      sheet.cell(CellIndex.indexByString("E$row")).value = TextCellValue(
          datesAndModes); // Date di collegamento + modalità senza duplicati

      row++;
    });

    // Salva il file in memoria
    final appDocumentsDir = await getApplicationDocumentsDirectory();
    String documentsPath =
        appDocumentsDir.path.replaceAll('Library/Containers', 'Documents');
    final outputPath = path.join(documentsPath, 'Presenze.xlsx');
    final outputFile = File(outputPath);
    await outputFile.writeAsBytes(await excel.encode()!);

    return outputFile;
  }
}
