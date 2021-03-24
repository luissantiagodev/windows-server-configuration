
const admin = require("firebase-admin");
const serviceAccount = require(`${__dirname}/../config/firebase.json`);

admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
    databaseURL: "https://mapfre-a072b-default-rtdb.firebaseio.com/",
    storageBucket: "gs://mapfre-a072b.appspot.com",
});

module.exports = {
    storage: admin.storage(),
    admin
};
