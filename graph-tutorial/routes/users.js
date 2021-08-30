var express = require('express');
var router = express.Router();

/* GET users listing. */
router.get('/', function (req, res, next) {
  if (!req.session.userId) {
    // Redirect unauthenticated requests to home page
    res.redirect('/')
  } else {
  }
  res.send('respond with a resource');
});

module.exports = router;
