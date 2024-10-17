const xlsx = require('xlsx');
const { MongoClient, ObjectId } = require('mongodb');

async function connectToMongo() {
  const uri = 'mongodb://127.0.0.1:27017/assets-app-test';
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
  const workbook = xlsx.readFile('../excel/employees.xlsx'); // Asegúrate de que la ruta al archivo Excel sea correcta
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rows = xlsx.utils.sheet_to_json(sheet);

  let updatedCount = 0;
  let notUpdatedCount = 0;
  const notUpdatedIds = [];

  for (const row of rows) {
    const assetIdStr = row['assetId'];
    const employeeIdStr = row['employeeId'];

    if (!assetIdStr || !employeeIdStr) {
      console.log(`Fila inválida en el Excel: ${JSON.stringify(row)}`);
      notUpdatedCount++;
      notUpdatedIds.push(assetIdStr);
      continue;
    }

    console.log(`Procesando activo ID: ${assetIdStr} para empleado ID: ${employeeIdStr}`);

    // Convertir los IDs a ObjectId
    let assetId, employeeId;
    try {
      assetId = new ObjectId(assetIdStr);
      employeeId = new ObjectId(employeeIdStr);
    } catch (error) {
      console.error(`Error al convertir los IDs: Activo ID ${assetIdStr}, Empleado ID ${employeeIdStr}. Error: ${error}`);
      notUpdatedCount++;
      notUpdatedIds.push(assetIdStr);
      continue;
    }

    try {
      // Retrieve asset details from assets collection
      const asset = await client.db('assets-app-test').collection('assets').findOne({ _id: assetId });
      if (!asset) {
        console.log(`Activo con ID ${assetIdStr} no encontrado. Verifica que el ID del activo sea correcto en la base de datos.`);
        notUpdatedCount++;
        notUpdatedIds.push(assetIdStr);
        continue;
      }

      // Retrieve employee details from employees collection
      const employee = await client.db('assets-app-test').collection('employees').findOne({ _id: employeeId });
      if (!employee) {
        console.log(`Empleado con ID ${employeeIdStr} no encontrado. Verifica que el ID del empleado sea correcto en la base de datos.`);
        notUpdatedCount++;
        notUpdatedIds.push(assetIdStr);
        continue;
      }

      // Update asset's assigned fields
      await client.db('assets-app-test').collection('assets').updateOne(
        { _id: assetId },
        {
          $set: {
            assigned: employee._id.toString(), // Asegurando que sea un string
            assignedTo: employee.name // Cambiado a assignedTo
          }
        }
      );
      console.log(`Se actualizó el activo ${assetIdStr} con el nuevo empleado ${employeeIdStr}`);

      // Add asset to new Employee
      const newAsset = {
        id: asset._id.toString(), // Asegurando que sea un string
        name: asset.name,
        brand: asset.brand,
        model: asset.model,
        assigned: false, // Asumiendo que "assigned" es un booleano
        EPC: asset.EPC,
        serial: asset.serial,
        creationDate: asset.creationDate, // Usando la fecha de creación original
      };

      await client.db('assets-app-test').collection('employees').updateOne(
        { _id: employeeId },
        { $push: { assetsAssigned: newAsset } }
      );
      console.log(`Se asignó el activo ${assetIdStr} al empleado ${employeeIdStr}`);

      updatedCount++;
    } catch (err) {
      console.error(`Error procesando activo ID: ${assetIdStr} para empleado ID: ${employeeIdStr}`, err);
      notUpdatedCount++;
      notUpdatedIds.push(assetIdStr);
    }
  }

  console.log(`Total de activos actualizados: ${updatedCount}`);
  console.log(`Total de activos no actualizados: ${notUpdatedCount}`);
  console.log(`IDs de activos no actualizados: ${notUpdatedIds.join(', ')}`);
}

// Llama a la función para conectar a MongoDB
connectToMongo();
