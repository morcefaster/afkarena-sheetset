'use strict';
const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const SCOPES = ['https://www.googleapis.com/auth/drive'];
const TOKEN_PATH = 'token.json';
const FOLDER_ID = {
    NN: '1X4ggCmEt_mQx-wYJI0Q8fQc1pooQYRlI',
    NNX: '1ILVV2i_oCgKG15N_R39ndoA7Irw0I4TB',
    NNZ: '1JBu30tMYImKs_fg-pAPDG_yYXqw51jrm'
}
const MAIN_SHEET_ID = {
    NN: '196UVWY6wyWcdBH7Ao-J4JTqqGJU2QGiAf2bwzKr54qA',
    NNX: '1__QoQj_q3ahwzrr8ul-39Fk_ZNIt4USlufCV8ZVU2L0',
    NNZ: '1Zi62Io2PeTnQk65ALKHCe2_jEpVQw-mwzQTXuaSOtkI'
}
const TEMPLATE_SHEET_ID = '1svlCIzO0B_0grjgcIket9AaB_GMUC5-2DemP78vub1I';

const HEROES_SHEET_PREFIX = "heroes_";
const HERO_COUNT = 110;
const SINGLE_SHEET_ID = 100;
const DROPDOWN_GRID_COORDINATES = {
    sheetId: 0,
    rowIndex: 0,
    columnIndex: 20
}

exports.handler = async (event) => {
    // TODO implement
    var response = {
        statusCode: 200,
        body: JSON.stringify('Hello from Lambda!'),
        headers: { "access-control-allow-origin": '*' }
    };

    console.log(`event:${JSON.stringify(event)}`)

    try {
        await upload_heroes(event.body)
            .then((ret) => {
                console.log(JSON.stringify(ret));
                response.body = JSON.stringify({ sheetUrl: ret.playerSheetUrl });
            });
    } catch (ex) {
        console.log(ex);
        response.statusCode = 500;
        response.body = JSON.stringify("An error has occurred. - " + JSON.stringify(ex));
    }
    console.log(`Response ${JSON.stringify(response)}`);
    return response;
};


async function upload_heroes(body) {
    return new Promise(function (resolve, reject) {        
        var parsed_body = JSON.parse(body);
        var parsed_heroes = JSON.parse(parsed_body.import);
        var req = {
            email: parsed_body.email,
            guild: parsed_body.guild.toUpperCase(),
            heroes: parsed_heroes[3].heroes,
            playerName: parsed_heroes[3].playerName,
            userId: parsed_heroes[3].ownerId,
            copiedOver: false
        }

        if (!req.playerName) {
            throw "Player name must be a part of JSON. Sign in to afkalc.";
        }

        if (!req.userId) {
            throw "User ID must be a part of JSON. Sign in to afkalc.";
        }

        if (!req.guild) {
            throw "Bad request";
        }

        var credentials = fs.readFileSync('credentials.json');
        req.auth = authorize(JSON.parse(credentials));
        console.log("getting sheets");
        req.sheets = google.sheets({ version: 'v4', auth: req.auth });
        req.drive = google.drive({ version: 'v3', auth: req.auth });
        var playerSheetUrl;
        readHeroNames(req)
            .then(findSheet)
            .then(copyTemplate)
            .then(grantPermissions)
            .then(getSheet)
            .then(updateSheet)
            .then((req) => {
                return new Promise(function(resolve, reject) {
                    playerSheetUrl = req.sheet.data.spreadsheetUrl;
                    req.sheetId = MAIN_SHEET_ID[req.guild];
                    return resolve(req);
                });
            })
            .then(getSheet)
            .then(updateMainSheet)
            .then(() => { return resolve({ playerSheetUrl }) })
            .catch((err) => {
                console.log("ERROR - " + JSON.stringify(err));
                return reject(err);
            });
    });
}

function authorize(credentials, callback) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirect_uris[0]);

    // Check if we have previously stored a token.
    console.log("authorizing with token");
    var token = fs.readFileSync(TOKEN_PATH);
    oAuth2Client.setCredentials(JSON.parse(token));
    console.log("authorized");
    return oAuth2Client;
}


async function readHeroNames(req) {
    return new Promise(function(resolve, reject) {
        fs.readFile('hero_names.json',
            (err, content) => {
                if (err) {
                    return reject(err);
                }
                req.heroNames = JSON.parse(content);
                return resolve(req);
            });
    });
}

async function updateMainSheet(req) {
    return new Promise(function (resolve, reject) {
        var requests = [];
        var heroSheet = req.sheet.data.sheets.find(s => last(s.properties.title.split('::')) === req.userId);
        var playerSheetId;
        if (heroSheet) {
            requests.push({
                updateSheetProperties: {
                    properties: {
                        sheetId: heroSheet.properties.sheetId,
                        title: `${HEROES_SHEET_PREFIX}${req.playerName}::${req.guild}::${req.userId}`,
                        hidden: true
                    },
                    fields: 'title, hidden'
                }
            });
            playerSheetId = heroSheet.properties.sheetId;
        } else {
            requests.push({
                addSheet: {
                    properties: {
                        title: `${HEROES_SHEET_PREFIX}${req.playerName}::${req.guild}::${req.userId}`,
                        sheetId: req.sheet.data.sheets.length + 1,
                        hidden: true
                    }
                }
            });
            playerSheetId = req.sheet.data.sheets.length+1;
        }

        requests.push({
            "updateCells": createUpdateCellsDataFromHeroes(req, playerSheetId)
        });

        var hero_sheet_names = req.sheet.data.sheets
            .filter(s => s.properties.title.startsWith(HEROES_SHEET_PREFIX))
            .map(s => s.properties.title.substring(HEROES_SHEET_PREFIX.length));

        if (!heroSheet) {
            hero_sheet_names.push(`${req.playerName}::${req.guild}::${req.userId}`);
        }

        var condition_values = hero_sheet_names.map(name => ({ userEnteredValue:name }));

        requests.push({
            "updateCells": {
                rows: [
                    {
                        values: [
                            {
                                userEnteredValue: {
                                    stringValue: hero_sheet_names[0]
                                },
                                dataValidation: {
                                    condition: {
                                        type: "ONE_OF_LIST",
                                        values: condition_values
                                    },
                                    showCustomUi: true
                                }
                            }
                        ]
                    }
                ],
                fields: "*",
                start: DROPDOWN_GRID_COORDINATES
            }
        });

        var batchUpdateRequest = { requests };

        req.sheets.spreadsheets.batchUpdate({
                spreadsheetId: req.sheet.data.spreadsheetId,
                resource: batchUpdateRequest,
                includeSpreadsheetInResponse: true
            },
            function (err, spreadsheet) {
                if (err) {
                    return reject(err);
                }
                return resolve(req);
            });
    });
}

async function grantPermissions(req) {
    return new Promise(function (resolve, reject) {
        if (!req.copied_over) {
            console.log("Permission grant not required");
            return resolve(req);
        }
        req.drive.permissions.create({
                resource: {
                    role: "writer",
                    type: "user",
                    emailAddress: req.email
                },
                fileId: req.sheetId,
                sendNotificationEmail: false
            },
            function(err, res) {
                if (err) {
                    console.log(err);
                    return reject(err);
                }
                console.log("Permissions granted");
                return resolve(req);
            });
    });
}

async function copyTemplate(req) {
    return new Promise(function (resolve, reject) {
        if (req.sheetId) {
            console.log("File exists, not copying");
            return resolve(req);
        }
        req.drive.files.copy({
                resource: {
                    name: `${req.playerName}::${req.guild}::${req.userId}`,
                    parents: [FOLDER_ID[req.guild]]
                },
                fileId: TEMPLATE_SHEET_ID
            },
            function(err, res) {
                if (err) {
                    return reject(err);
                }
                console.log("Copied the template over");
                req.copied_over = true;
                req.sheetId = res.data.id;
                return resolve(req);
            });
    });
}



/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES
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
            callback(oAuth2Client);
        });
    });
}

async function getSheet(req) {
    return new Promise(function(resolve, reject) {
        req.sheets.spreadsheets.get({
            spreadsheetId: req.sheetId
            },
            (err, res) => {
                if (err) {
                    console.log('The API returned an error: ' + err);
                    console.log("Sheet did not get copied?");
                    return reject(err);
                }
                console.log("Retrieved sheet " + req.sheetId);
                req.sheet = res;
                return resolve(req);
            });
    });
}


async function updateSheet(req) {
    return new Promise(function(resolve, reject) {
        var requests = [];

        requests.push({
            updateSpreadsheetProperties: {
                properties: {
                    title: `${req.playerName}::${req.guild}::${req.userId}`
                },
                fields: 'title'
            }
        });

        var heroSheet = req.sheet.data.sheets.find(s => last(s.properties.title.split('::')) === req.userId);
        
        var singleUserSheetId;
        if (heroSheet) {
            requests.push({
                updateSheetProperties: {
                    properties: {
                        sheetId: heroSheet.properties.sheetId,
                        title: `${HEROES_SHEET_PREFIX}${req.playerName}::${req.guild}::${req.userId}`,
                        hidden: true
                    },
                    fields: 'title,hidden'
                }
            }); 
            singleUserSheetId = heroSheet.properties.sheetId;
        } else {
            requests.push({
                addSheet: {
                    properties: {
                        title: `${HEROES_SHEET_PREFIX}${req.playerName}::${req.guild}::${req.userId}`,
                        sheetId: SINGLE_SHEET_ID,
                        hidden: true
                    }
                }
            });
            singleUserSheetId = SINGLE_SHEET_ID;
        }

        requests.push({
            "updateCells": createUpdateCellsDataFromHeroes(req, singleUserSheetId)
        });

        requests.push({
            "updateCells": {
                rows: [
                    {
                        values: [
                            {
                                userEnteredValue: {
                                    stringValue: `${req.playerName}::${req.guild}::${req.userId}`
                                }
                            }
                        ]
                    }
                ],
                fields: "*",
                start: DROPDOWN_GRID_COORDINATES
            }
        });

        var batchUpdateRequest = { requests };

        req.sheets.spreadsheets.batchUpdate({
                spreadsheetId: req.sheet.data.spreadsheetId,
                resource: batchUpdateRequest
            },
            function(err, res) {
                if (err) {
                    return reject(err);
                }
                return resolve(req);
            });
    });
}

function createUpdateCellsDataFromHeroes(req, sheetId) {
    var rowData = [];
    var data = {
        start: {
            sheetId: sheetId,
            columnIndex: 0,
            rowIndex: 0
        },
        fields: "*"
    }

    var headerRow = sa2row(["hero_name", "hero_id", "ascension", "signature item", "furniture", "engraving"]);
    rowData.push(headerRow);
    for (var hero_id = 1; hero_id < HERO_COUNT; hero_id++) {
        var h = req.heroes[hero_id];
        var h_name = req.heroNames[hero_id];
        if (h_name) {
            rowData.push(sa2row([
                `${h_name}`,
                `${hero_id}`,
                `${h ? (h.ascend) : 0}`,
                `${h ? (typeof h.si === 'undefined' ? -1 : h.si) : -1}`,
                `${h ? (h.fi || 0) : 0}`,
                `${h ? (h.engrave || 0) : 0}`
            ]));
        }
    }
    data.rows = rowData;
    return data;
}


async function findSheet(req) {
    return new Promise(function(resolve, reject) {
        req.drive.files.list({
            q: `parents in '${FOLDER_ID[req.guild]}'`,
            fields: 'files(id,name,trashed)'
        },
            function(err, res) {
                if (err) {
                    return reject(err);
                }
                var my_sheet_file = res.data.files.find(f => last(f.name.split('::')) === req.userId && !f.trashed);
                if (my_sheet_file) {
                    req.sheetId = my_sheet_file.id;
                }
                return resolve(req);
            });
    });
}


function sa2row(arr) {
    var val_arr = [];
    arr.forEach(a => val_arr.push({ userEnteredValue: { stringValue: a } }));
    return { values: val_arr };
}

function last(arr) {
    return arr[arr.length - 1];
}