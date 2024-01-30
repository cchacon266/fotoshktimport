const xlsx = require('xlsx');
const { MongoClient, ObjectId } = require('mongodb');

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
    const activoId = new ObjectId(row.id);

    // Obtén los valores de las columnas del Excel
    const valAdqUSD1 = row['Val.adq. USD 1'];
    const valContMN1 = row[' Val.cont. M.N. 1'];

    const activo = await client.db('assets-app-grupomexico').collection('assets').findOne({ _id: activoId });

    if (activo) {
      // Busca el campo existente y actualiza su valor
      actualizarCampoTexto(activo, 'Val.adq. USD 1', valAdqUSD1);
      actualizarCampoTexto(activo, ' Val.cont. M.N. 1', valContMN1);

      // Actualiza la base de datos con la nueva información
      await client.db('assets-app-grupomexico').collection('assets').updateOne(
        { _id: activoId },
        { $set: { customFieldsTab: activo.customFieldsTab } }
      );

      console.log(`Se actualizó el activo con ID ${activoId}`);
    } else {
      console.log(`No se encontró el activo con ID ${activoId}`);
    }
  }
}

function actualizarCampoTexto(activo, fieldName, value) {
  const existingIndex = activo.customFieldsTab['tab-0'].left.findIndex(item => item.values.fieldName === fieldName);

  if (existingIndex !== -1) {
    // Si existe, actualizar el valor del campo existente
    activo.customFieldsTab['tab-0'].left[existingIndex].values.initialValue = value !== undefined ? Number(value).toFixed(2) : '';
  } else {
    // Si no existe, agregar un nuevo campo para el nuevo valor
    activo.customFieldsTab['tab-0'].left.push({
      id: `${fieldName}-${activo._id}`,
      content: 'singleLine',
      values: {
        fieldName,
        initialValue: value !== undefined ? Number(value).toFixed(2) : ''
      }
    });
  }
}

// Llama a la función para conectar a MongoDB
connectToMongo();
