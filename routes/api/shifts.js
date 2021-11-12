const express = require("express");
const router = express.Router();
const auth = require("../../middleware/auth");
const { body, validationResult } = require("express-validator");
const moment = require('moment');
const Shift = require("../../models/Shift");

// @route       POST api/shifts
// @desc        Create or Update Shift
// @access      Private
router.post(
  "/",
  auth,
  [
    body("shiftName", "Shift name is required").not().isEmpty(),
  ],
  async (req, res) => {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).json({ errors: errors.array() });
    }

    const { id, shiftName, timeStart, timeEnd, branchId } = req.body;
    
    // Xét giá trị là ngày thứ 2 của tuần sau
    var date = moment().startOf("isoWeek").add(7, "days").format('YYYY-MM-DD');
    var shiftTime = null;
    var time = null;

    let hours1 = null;
    let hours2 = null;
    let minute1 = null;
    let minute2 = null;
    let subMinute = null;
    let startTime = timeStart;
    let endTime = timeEnd;

    if (startTime.indexOf(":") === 1) startTime = "0" + startTime;
    if (endTime.indexOf(":") === 1) endTime = "0" + endTime;
    time = startTime + "-" + endTime
    hours1 = startTime.substring(0, 2);
    hours2 = endTime.substring(0, 2);
    minute1 = startTime.substring(3, 5);
    minute2 = endTime.substring(3, 5);

    if (minute2 > minute1) {
      subMinute = minute2 - minute1;
    } else {
      subMinute = minute1 - minute2;
    }

    shiftTime = hours2 - hours1;
    subMinute = subMinute / 60;
    shiftTime = (Number(shiftTime) + Number(subMinute)).toString();

    // Kiểm tra xem đã có record hay chưa
    // Mục đích để tăng giá trị cho biến position
    let position = 0;
    let shift1 = await Shift.findOne({
      $and: [{ branchId: branchId }, { date: date }],
    }).sort({ date: -1, position: -1 });
    if (shift1) {
      position = shift1.position + 1;
    }

    const shiftField = {};
    if (id) shiftField.id = id;
    shiftField.shiftName = shiftName;
    shiftField.shiftTime = shiftTime;
    shiftField.time = time;
    shiftField.branchId = branchId;
    shiftField.date = date;

    try {
      let shift = await Shift.findOne({
        $and: [{ shiftName: shiftName }, { date: date }],
      });

      if (id) {
        // Update
        shift = await Shift.findOneAndUpdate(
          { _id: id },
          { $set: shiftField },
          { new: true }
        );
        return res.json(shift);
      }

      if (shift) {
        return res
          .status(400)
          .json({ errors: [{ msg: "Shift already exists" }] });
      }

      // Create
      shift = new Shift({
        shiftName,
        shiftTime,
        time,
        branchId,
        position,
        date,
      });

      await shift.save();

      res.json(shift);
    } catch (err) {
      console.error(err.message);
      res.status(500).send("Server error");
    }
  }
);

// @route       GET api/shift
// @desc        Get all Shifts
// @access      Public
router.get("/", async (req, res) => {
  try {
    const shifts = await Shift.find().populate("branchId").sort({ date: -1 });
    res.json(shifts);
  } catch (err) {
    console.error(err.message);
    res.status(500).send("Server Error");
  }
});

// @route       GET api/shifts/:id
// @desc        Get Shift by ID
// @access      Private
router.get("/:id", auth, async (req, res) => {
  try {
    console.log("server 1");
    const shift = await Shift.findById(req.params.id);

    if (!shift) {
      return res.status(404).json({ msg: "Shift not found" });
    }

    res.json(shift);
  } catch (err) {
    console.error(err.message);
    if (err.kind === "ObjectId") {
      return res.status(404).json({ msg: "Shift not found" });
    }
    res.status(500).send("Server Error");
  }
});

// @route       GET api/shifts/:branchId/:dateFrom
// @desc        Get Shift by branchId, dateFrom
// @access      Private
router.get("/:branchId/:dateFrom", auth, async (req, res) => {
  try {
    const shift = await Shift.find({
      $and: [{ branchId: req.params.branchId }, { date: { $lte: req.params.dateFrom } }],
    }).sort({ date: 1 });

    if (!shift) {
      return res.status(404).json({ msg: "Shift not found" });
    }

    var getDateFinal = "";
    shift.map((ele) => {
      getDateFinal = ele.date;
      if (getDateFinal > ele.date) {
        getDateFinal = ele.date;
      }
    })

    const shift1 = await Shift.find({ date: getDateFinal }).sort({ shiftName: 1 });

    res.json(shift1);
  } catch (err) {
    console.error(err.message);
    if (err.kind === "ObjectId") {
      return res.status(404).json({ msg: "Shift not found" });
    }
    res.status(500).send("Server Error");
  }
});

// @route       GET api/shifts/shiftForSalary/:dateFrom/:dateTo
// @desc        Get Shift by dateFrom, dateTo
// @access      Private
router.get("/shiftForSalary/:dateFrom/:dateTo", auth, async (req, res) => {
  try {
    const shift = await Shift.find().sort({ date: -1 });

    if (!shift) {
      return res.status(404).json({ msg: "Shift not found" });
    }

    var getStartDate = "";
    shift.map((ele) => {
      if (moment(ele.date).format("YYYY-MM-DD") <= req.params.dateFrom && getStartDate === "") {
        getStartDate = moment(ele.date).format("YYYY-MM-DD");
      }
    })

    const shift1 = await Shift.find({
      $and: [{ date: { $gte: getStartDate } }, { date: { $lte: req.params.dateTo } }],
    }).sort({ date: -1 });

    res.json(shift1);
  } catch (err) {
    console.error(err.message);
    if (err.kind === "ObjectId") {
      return res.status(404).json({ msg: "Shift not found" });
    }
    res.status(500).send("Server Error");
  }
});

// @route       DELETE api/shifts/:id
// @desc        Delete a Shift
// @access      Private
router.delete("/:id", auth, async (req, res) => {
  try {
    const shift = await Shift.findById(req.params.id);

    if (!shift) {
      return res.status(404).json({ msg: "Shift not found" });
    }

    await shift.remove();

    res.json({ msg: "Shift removed" });
  } catch (err) {
    if (err.kind === "ObjectId") {
      return res.status(404).json({ msg: "Shift not found" });
    }
    res.status(500).send("Server Error");
  }
});

module.exports = router;
