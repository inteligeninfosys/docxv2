const express = require('express');
const app = express();
const router = express.Router();
const cors = require('cors')

//include the routes file
var repossessionsendphy_word = require('./routes/repossessionsendphy_word');
var repossessionsendphy_wordv1 = require('./routes/repossessionsendphy_wordv1');
var repossessionsendphy_wordv1pdf = require('./routes/repossessionsendphy_wordv1pdf');
var demand2 = require('./routes/demand2');
var demand1 = require('./routes/demand1');

app.use('/docxv2/repossessionsendphy_word', repossessionsendphy_word);
app.use('/docxv2/repossessionsendphy_wordv1', repossessionsendphy_wordv1);
app.use('/docxv2/repossessionsendphy_wordv1pdf', repossessionsendphy_wordv1pdf);
app.use('/docxv2/demand2', demand2);
app.use('/docxv2/demand1', demand1);

router.get('/', function (req, res) { 
  res.json({ message: 'Demand letters ready Home!' });
});

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.use(cors())

//add the router
app.use('/docxv2', router);
app.listen(process.env.port || 8040);

console.log('Running at Port 8040');