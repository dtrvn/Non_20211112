import React, { Fragment, useState, useEffect } from "react";
import { Link, withRouter } from "react-router-dom";
import PropTypes from "prop-types";
import { connect } from "react-redux";
import { addShift } from "../../actions/shift";
import TimeKeeper from 'react-timekeeper';

const AddShiftForm = ({ addShift, branchs, history }) => {
  const initialFormState = { id: null, shiftName: "", shiftTime: "", timeStart: "09:00", timeEnd: "12:00", date: "", branchId: null };

  const [formData, setFormData] = useState(initialFormState);

  const [showTimeStart, setShowTimeStart] = useState(false);
  const [showTimeEnd, setShowTimeEnd] = useState(false);

  const onChange = (e) =>
    setFormData({ ...formData, [e.target.name]: e.target.value });

  const clearForm = () => {
    setFormData({ ...formData, shiftName: "", shiftTime: "", timeStart: "09:00", timeEnd: "12:00", date: "" });
  };

  useEffect(() => {
    if (branchs.length > 0) {
      setFormData({
        ...formData,
        branchId: branchs[0]._id,
      })
    }
  }, [branchs]);

  let elmBranchs = branchs.map((ele) => (
    <option value={ele._id}>{ele.branchAddress}</option>
  ));

  return (
    <Fragment>
      <div className="row">
        <form
          onSubmit={(e) => {
            e.preventDefault();
            addShift(formData, history);
            clearForm();
          }}
        >
          <div className="form-body">
            <div className="row p-t-5 m-l-5">
              <div className="col-md-addShift-2">
                <div className="form-group has-success">
                  <input type="text"
                    name="shiftName"
                    value={formData.shiftName}
                    onChange={(e) => onChange(e)}
                    className="form-control"
                    placeholder="* Tên ca"
                    required />
                </div>
              </div>
              <div className="col-md-addShift-2">
                <div className="form-group has-success">
                  <input type="text"
                    name="time"
                    value={formData.timeStart}
                    className="form-control"
                    placeholder="* Thời gian bắt đầu"
                    onChange={(e) => onChange(e)}
                    onClick={() => {
                      setShowTimeStart(!showTimeStart);
                      setShowTimeEnd(false);
                    }}
                    required />
                </div>
              </div>
              <div className="col-md-addShift-2">
                <div className="form-group has-success">
                  <input type="text"
                    name="time"
                    value={formData.timeEnd}
                    className="form-control"
                    placeholder="* Thời gian kết thúc"
                    onChange={(e) => onChange(e)}
                    onClick={() => {
                      setShowTimeEnd(!showTimeEnd);
                      setShowTimeStart(false);
                    }}
                    required />
                </div>
              </div>
              <div className="col-md-addShift-2">
                <div className="form-group has-success">
                  <select
                    name="branchId"
                    value={formData.branchId}
                    onChange={(e) => onChange(e)}
                    class="form-control custom-select"
                  >
                    {elmBranchs}
                  </select>
                </div>
              </div>
              <div className="col-md-addShift-2">
                <div className="form-actions">
                  <button type="submit" class="btn btn-info"><i class="fas fa-plus"></i> Thêm</button>
                </div>
              </div>
            </div>
          </div>
          {showTimeStart &&
            <TimeKeeper
              time={formData.timeStart}
              onChange={(newTime) => setFormData({
                ...formData,
                timeStart: newTime.formatted24
              })}
              hour24Mode
              onDoneClick={() => setShowTimeStart(false)}
              switchToMinuteOnHourSelect
            />
          }
          {showTimeEnd &&
            <TimeKeeper
              time={formData.timeEnd}
              onChange={(newTime) => setFormData({
                ...formData,
                timeEnd: newTime.formatted24
              })}
              hour24Mode
              onDoneClick={() => setShowTimeEnd(false)}
              switchToMinuteOnHourSelect
            />
          }
        </form>
      </div>
    </Fragment>
  );
};

AddShiftForm.propTypes = {
  addShift: PropTypes.func.isRequired,
  branchs: PropTypes.object.isRequired,
};

export default connect(null, { addShift })(withRouter(AddShiftForm));
