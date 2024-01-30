const xlsx = require('xlsx');
const sharp = require('sharp');
const fs = require('fs');
const { MongoClient, ObjectId } = require('mongodb');

// Programa con fotos reducidas

async function connectToMongo() {
  const uri = 'mongodb://127.0.0.1:27017/assets-app-grupomexico';
  const client = new MongoClient(uri, { useNewUrlParser: true, useUnifiedTopology: true });

  try {
    await client.connect();
    console.log('Conectado a MongoDB');

    // Llamamos a la función principal del programa
    await main(client);
  } catch (error) {
    console.error('Error al conectar a MongoDB:', error);
  } finally {
    await client.close();
    console.log('Desconectado de MongoDB');
  }
}

async function main(client) {
  const workbook = xlsx.readFile('../excel/assetstotal.xlsx');
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rows = xlsx.utils.sheet_to_json(sheet);
  for (const row of rows) {
    // Verificar si row.id es un valor válido para ObjectId
    if (!row.id || row.id.length !== 24) {
      console.warn(`No se pudo convertir a ObjectId: ${row.id}`);
      console.log(`No se pudo convertir 'row.id' a un formato aceptable para ObjectId: ${row.id}`);
      continue;  // Saltar a la siguiente iteración del bucle
    }

    const activoId = new ObjectId(row.id);
    const activo = await client.db('assets-app-grupomexico').collection('assets').findOne({ _id: activoId });

    if (activo) {
      for (let i = 1; i <= 6; i++) {
        const fieldName = `Imagen ${i}`;
        const fileName = row[fieldName];

        if (fileName) {
          const [name, extension] = fileName.split('.') || [fileName];
          const formattedExtension = extension ? extension.toLowerCase() : '';

          // Verificar si ya existe un campo con el mismo fieldName
          const existingIndex = activo.customFieldsTab['tab-0'].left.findIndex(item => item.values.fieldName === fieldName);

          if (existingIndex !== -1) {
            // Si existe, actualizar las nuevas imágenes al campo existente
            activo.customFieldsTab['tab-0'].left[existingIndex].values.fileName = name;
            activo.customFieldsTab['tab-0'].left[existingIndex].values.initialValue = formattedExtension;
          } else {
            // Si no existe, agregar un nuevo campo para las nuevas imágenes
            activo.customFieldsTab['tab-0'].left.push({
              id: `imagen-${i}-${activoId}`,
              content: 'imageUpload',
              values: {
                fieldName,
                initialValue: formattedExtension,
                fileName: name
              }
            });
          }

          // Reducir el peso de las imágenes y guardarlas
          const imagenPath = `../fotosgm/${fileName}`;
          const imagenReducidaPath = `fotosgm/customFields/${name}.${formattedExtension}`;

          // Verificar si el archivo existe antes de copiarlo
          if (fs.existsSync(imagenPath)) {
            // Reducir el peso de la imagen y guardarla con calidad del 60%
            await sharp(imagenPath).jpeg({ quality: 60 }).toFile(imagenReducidaPath);
            console.log(`Imagen reducida y guardada: ${imagenReducidaPath}`);
          } else {
            console.warn(`La imagen no existe: ${imagenPath}`);
          }
        }
      }

      // Actualizar la base de datos con la nueva información
      await client.db('assets-app-test').collection('assets').updateOne(
        { _id: activoId },
        { $set: { 'customFieldsTab': activo.customFieldsTab } }
      );

      console.log(`Se actualizó el activo con ID ${activoId}`);
    } else {
      console.log(`No se encontró el activo con ID ${activoId}`);
    }
  }
}

// Llama a la función para conectar a MongoDB
connectToMongo();
