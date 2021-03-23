const firebaseHelper = require('./helpers/FirebaseHelper')


firebaseHelper.admin.database().ref("listener_documents").on('value', (snapshot) => {
    console.log(snapshot.val())
})