const router = require('express-promise-router')();
const graph = require('../graph.js');
const addDays = require('date-fns/addDays');
const formatISO = require('date-fns/formatISO');
const formatDate = require('date-fns/format');
const startOfWeek = require('date-fns/startOfWeek');
const zonedTimeToUtc = require('date-fns-tz/zonedTimeToUtc');
const iana = require('windows-iana');
const { body, validationResult } = require('express-validator');
const validator = require('validator');

/* GET /calendar */
router.get('/',
  async function (req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      const params = {
        active: { calendar: true }
      };

      // Get the user
      const user = req.app.locals.users[req.session.userId];
      // Convert user's Windows time zone ("Pacific Standard Time")
      // to IANA format ("America/Los_Angeles")
      const timeZoneId = iana.findIana(user.timeZone)[0];
      console.log(`Time zone: ${timeZoneId.valueOf()}`);

      // Calculate the start and end of the current week
      // Get midnight on the start of the current week in the user's timezone,
      // but in UTC. For example, for Pacific Standard Time, the time value would be
      // 07:00:00Z
      var weekStart = zonedTimeToUtc(startOfWeek(new Date()), timeZoneId.valueOf());
      var weekEnd = addDays(weekStart, 7);
      console.log(`Start: ${formatDate(weekStart, "yyyy-MM-dd'T'HH:mm:SS'Z'")}`);
      //console.log(`Start: ${formatISO(weekStart)}`);

      try {
        // Get the events
        const events = await graph.getCalendarView(
          req.app.locals.msalClient,
          req.session.userId,
          formatDate(weekStart, "yyyy-MM-dd'T'HH:mm:SS'Z'"),
          formatDate(weekEnd, "yyyy-MM-dd'T'HH:mm:SS'Z'"),
          user.timeZone);

        // Assign the events to the view parameters
        params.events = events.value;
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not fetch events',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
      }
      res.render('calendar', params);
    }
  }
);

module.exports = router;