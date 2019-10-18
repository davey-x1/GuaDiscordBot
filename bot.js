// Constantes de discord
const Discord = require('discord.js');
const client = new Discord.Client();
// Constantes de excel

const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets']; // If modifying these scopes, delete token.json.
const TOKEN_PATH = 'token.json';
var oAuth2Client;
var actualMessage;



// Starting functions

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  // Authorize a client with credentials, then call the Google Sheets API.
  authorize(JSON.parse(content));
});


// ------------------------------------------------------------------------------------------- //
// -------------------------------- Discord Logic ------------------------------------------- //
// ------------------------------------------------------------------------------------------- //

client.on('ready', () => {
  console.log(`Logged in as ${client.user.tag}!`);
});

client.on('message', msg => {
	actualMessage = msg;
	var msgArray = msg.content.split(' ');
	switch(msgArray[0]){
		// ---------- for tests only -------------- //
		case 'stop':
			msg.reply('Good bye :)');
			throw new FatalError("Something went badly wrong!");
		break;
		// --------- case Hi ---------- ///
		case 'hi':
			msg.reply('Hello!');
		break;
		// ----------- case Join ---------- //
		case 'join':
			var voiceChannel = msg.member.voiceChannel;
			if(!voiceChannel){
				msg.reply('No te encuentro en ningún canal de voz');
				break;
			} else {
				voiceChannel.join()
				.then(connection => console.log('Connected!'))
				.catch(console.error);
			}
		break;
		// --------------- Balance logic  --------------- //
		case '!balance':
			if(!msgArray[1]){
				msg.reply('El comando es !balance <nombre del jugador> <EUR/GOLD>, te falta añadir el nombre, usa !sellers para ver'
				+ ' los jugadores registrados');
				break;
			} else {
				if(!msgArray[2]){
					msg.reply('El comando es !balance <nombre del jugador> <EUR/GOLD>, te falta añadir el tipo de cambio.');
				} else {
					checkBalance(oAuth2Client, msgArray[1], msgArray[2]);
				}
				
			}
			
		break;
		// ------------ List of sellers ---------- //
		case '!sellers':
			returnSellers(oAuth2Client);
		break;
		// ------------- Total of Month --------------- //
		case '!pot':
			if(msgArray[1] == 'EUR'){
				returnPot(oAuth2Client, 1);
				break;
			} else if (msgArray[1] == 'GOLD'){
				returnPot(oAuth2Client, 2);
				break;
			} else {
				msg.reply('El comando correcto es !pot <EUR/GOLD');
			}
		break;
		// ------------- Add/Remove from balance --------- //
		case '!change':
			if(!msgArray[1]){
				msg.reply('The command: !change <remove/add> <name> <EUR/GOLD> <quantity>, missing <remove/add> function');
				break;
			} else {
				if(!msgArray[2]){
					msg.reply('The command: !change <remove/add> <name> <EUR/GOLD> <quantity>, missing <name>, you can check seller names using !sellers');
				} else {
					if(!msgArray[3] || (msgArray[3] != 'GOLD' && msgArray[3] != 'EUR')){
						msg.reply('The command: !change <remove/add> <name> <EUR/GOLD> <quantity>, missing <EUR/GOLD> option or selected a wrong option');
					} else {
						if(!msgArray[4]){
							msg.reply('The command: !change <remove/add> <name> <EUR/GOLD> <quantity>, missing <quantity>');
						} else {
							setMoney(oAuth2Client, msgArray[1], msgArray[2], msgArray[3], msgArray[4]);
						}
					}
				}
				
			}
		break;
		// ---------------------------------------------- //
		case '!add':
			setMoney(oAuth2Client, 'add', 'Jota', 'EUR', msgArray[1]);
			setMoney(oAuth2Client, 'add', 'Davey', 'EUR', msgArray[1]);
			setMoney(oAuth2Client, 'add', 'Polonia', 'EUR', msgArray[1]);
			setMoney(oAuth2Client, 'add', 'Idank', 'EUR', msgArray[1]);
		break;
	}
});

function checkType(stringUsed){
	if(stringUsed == 'EUR'){
		return true;
	} else if(stringUsed == 'GOLD') {
		return true;
	} else {
		return false;
	}
}

// ------------------------- Entrar a un canal ---------------------------- //
client.login('NjMyNDI2NzQ0MTQ5OTY2ODY4.XaFQbA.sD-CWxD26pSSangY7ECRkInw8yc');

// ------------------------------------------------------------------------

function FatalError(){
	Error.apply(this, arguments); this.name = "FatalError"; 
}
FatalError.prototype = Object.create(Error.prototype);



/* ------------------------------------------------------------------------------------------------------------------- 
 ---------------------------------------------- EXCEL -------------------------------------------------------------- 
 ------------------------------------------------------------------------------------------------------------------- */




/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client);
    oAuth2Client.setCredentials(JSON.parse(token));
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error while trying to retrieve access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
    });
  });
}

// ----------------------------------------------------------

function checkBalance(auth, name, tipo){
	var indexNombre;
	var indexMoneda;
	if(tipo == 'EUR'){
		indexMoneda = 1;
	} else {
		indexMoneda = 2;
	}
	// ------------------
	const sheets = google.sheets({version: 'v4', auth});
	sheets.spreadsheets.values.get({
		spreadsheetId: '18NJjtApSwFUCJkkbVmm-XteDW_M-BwkT86i2FB8Ut88',
		range: 'Octubre!D3:G',
	}, 
	(err, res) => {
		if(err){
			return console.log('La API lanzo un error' + err);
		}
		
		const rows = res.data.values;
		for(var index = 0; index < rows[0].length; index++){
			if(rows[0][index] == name){
				indexNombre = index;
				actualMessage.reply('El balance de ' + name + ' del mes actual en ' + tipo +' es: ' + rows[indexMoneda][indexNombre]);
				break;
			}
			
			if(index == rows[0].length - 1 && rows[0][index] != name){
				actualMessage.reply('The <name> is wrong, please check the registered users with the cmd !sellers');
				break;
			}
		}
	});
}

/*
	In short, return the active sellers
	-----------------------------------
*/
function returnSellers(auth){
	
	const sheets = google.sheets({version: 'v4', auth});
	sheets.spreadsheets.values.get({
		spreadsheetId: '18NJjtApSwFUCJkkbVmm-XteDW_M-BwkT86i2FB8Ut88',
		range: 'Octubre!D3:G',
	}, 
	(err, res) => {
		if(err){
			return console.log('La API lanzo un error' + err);
		}
		const rows = res.data.values;
		actualMessage.reply('Los nombres registrados son: ' + rows[0]);
	});
}

/*
	Add-Remove money from balance, need an auth object for excel connection
	change variable is for choose the function (Add or remove), team is the name of the seller
	type referencing for money and quantity is the value to process

*/

function setMoney(auth, change, team, type, quantity){
	var indexType;
	if(type == 'EUR'){
		indexType = 1;
	} else {
		indexType = 2;
	}
	
	const sheets = google.sheets({version: 'v4', auth});
	sheets.spreadsheets.values.get({
		spreadsheetId: '18NJjtApSwFUCJkkbVmm-XteDW_M-BwkT86i2FB8Ut88',
		range: 'Octubre!D3:G',
	}, 
	(err, res) => {
		if(err){
			return console.log('La API lanzo un error' + err);
		}
		const rows = res.data.values;
		for(var index = 0; index < rows[0].length; index++){
			if(rows[0][index] == team){
				if(change == 'remove'){
					changeMoney(auth, team, index, indexType, quantity, rows[indexType][index], 0);
				} else if (change == 'add') {
					changeMoney(auth, team, index, indexType, quantity, rows[indexType][index], 1);
				}
				break;
			}
			
			if(index == rows[0].length - 1 && rows[0][index] != name){
				actualMessage.reply('The <name> is wrong, please check the registered users with the cmd !sellers');
				break;
			}
		}
	});
	
	
}

/*
	In short, return the total amount of eur/gold won in a month
	------------------------------------------------------------
	
*/

function returnPot(auth, type){
	const sheets = google.sheets({version: 'v4', auth});
	sheets.spreadsheets.values.get({
		spreadsheetId: '18NJjtApSwFUCJkkbVmm-XteDW_M-BwkT86i2FB8Ut88',
		range: 'Octubre!D3:G',
	}, 
	(err, res) => {
		if(err){
			return console.log('La API lanzo un error' + err);
		}
		const rows = res.data.values;
		var totalValue = 0;
		for(var index = 0; index < rows[type].length; index++){
			totalValue = totalValue + Number(rows[type][index]);
		}
		
		actualMessage.reply('Total ganado este mes: ' + totalValue);
	});
}

/*
	Auxiliar function for setMoney()
	requires name, columna/fila for cells
	quantity is the amount to remove/add
	valorActual is the current balance of the seller
	tipo EUR/GOLD
	
*/

function changeMoney(auth, name, columna, fila, quantity, valorActual, tipo){
	// variables
	var columnaUsar;
	var filaUsar;
	var inputValue;
	var moneda;
	
	// Represent the tipo var to choose the row
	
	if(tipo == 0){
		inputValue = Number(valorActual) - quantity;
	} else {
		inputValue = Number(valorActual) + Number(quantity);
	}
	
	// Set the column for cell
			
	switch(columna){
		case 0:
			columnaUsar = 'D';
		break;
		case 1:
			columnaUsar = 'E';
		break;
		case 2:
			columnaUsar = 'F';
		break;
		case 3:
			columnaUsar = 'G';
		break;
	}
	
	// Set the row cell
	
	switch(fila){
		case 0:
			filaUsar = 3;
		break;
		case 1:
			moneda = 'EUR';
			filaUsar = 4;
		break;
		case 2:
			moneda = 'GOLD';
			filaUsar = 5;
		break;
	}
	
	// variables
	var row = columnaUsar + filaUsar;
	let values = [[inputValue]];
	let resource = {
		values,
	};
	// spreadsheet sheet
	let spreadsheetId = "18NJjtApSwFUCJkkbVmm-XteDW_M-BwkT86i2FB8Ut88";
	let range = "Octubre!"+row;
	let valueInputOption = "RAW";
	let myValue = 5;
	
	// excel process
	const sheets = google.sheets({version: 'v4', auth});
	sheets.spreadsheets.values.update({
	  spreadsheetId,
	  range,
	  valueInputOption,
	  includeValuesInResponse: true,
	  resource,
	}, (err, result) => {
	  if (err) {
		// Handle error
		console.log(err);
	  } else {
		actualMessage.reply("El nuevo balance de " + name + " es: " + result.data.updatedData.values);
	  }
	});
}