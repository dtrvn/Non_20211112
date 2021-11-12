import React, { Fragment, useState, useEffect } from "react";
import { Link, withRouter } from "react-router-dom";
import PropTypes from "prop-types";
import { connect } from "react-redux";
import { addTypeUser } from "../../actions/typeUser";

const AddTypeUserForm = ({
  addTypeUser,
  editing,
  currentTypeUser,
  history,
}) => {
  const initialFormState = {
    id: null,
    typeUsername: "",
    typeUserPercentCost: "",
  };

  const [formData, setFormData] = useState(initialFormState);

  const onChange = (e) =>
    setFormData({ ...formData, [e.target.name]: e.target.value });

  const clearForm = () => {
    setFormData({ ...formData, typeUsername: "", typeUserPercentCost: "" });
  };

  return (
    <Fragment>
      {/* <form
        class="form"
        onSubmit={(e) => {
          e.preventDefault();
          addTypeUser(formData, history);
          clearForm();
        }}
      >
        <div className="form-group add-flex" >
          <div class="col-md-4">
            <input
              type="text"
              placeholder="* Loại nhân viên"
              name="typeUsername"
              value={formData.typeUsername}
              onChange={(e) => onChange(e)}
              style={{ fontSize: '14px' }}
              required
            />

          </div>
          <div class="col-md-4" >
            <input
              type="text"
              placeholder="* Phần trăm lương"
              name="typeUserPercentCost"
              value={formData.typeUserPercentCost}
              onChange={(e) => onChange(e)}
              style={{ fontSize: '14px' }}
              required
            />
          </div>
          <div class="col-md-6" >
            <button type="submit" class="btn btn-info"><i class="fas fa-plus"></i>
              <span className="hide-sm"> Thêm</span>
            </button>
          </div>
        </div>
      </form> */}
      <div className="row">
        <form
          onSubmit={(e) => {
            e.preventDefault();
            addTypeUser(formData, history);
            clearForm();
          }}
        >
          <div className="form-body">
            <div className="row p-t-5 m-l-5">
              <div className="col-md-addJob-4">
                <div className="form-group has-success">
                  <input
                    type="text"
                    placeholder="* Loại nhân viên"
                    name="typeUsername"
                    className="form-control"
                    value={formData.typeUsername}
                    onChange={(e) => onChange(e)}
                    style={{ fontSize: '14px' }}
                    required
                  />
                </div>
              </div>
              <div className="col-md-addJob-4">
                <div className="form-group has-success">
                  <input
                    type="text"
                    placeholder="* Phần trăm lương"
                    name="typeUserPercentCost"
                    className="form-control"
                    value={formData.typeUserPercentCost}
                    onChange={(e) => onChange(e)}
                    style={{ fontSize: '14px' }}
                    required
                  />
                </div>
              </div>
              <div className="col-md-addJob-4">
                <div className="form-actions">
                  <button type="submit" class="btn btn-info"><i class="fas fa-plus"></i> Thêm</button>
                </div>
              </div>
            </div>
          </div>
        </form>
      </div>
    </Fragment>
  );
};

AddTypeUserForm.propTypes = {
  addTypeUser: PropTypes.func.isRequired,
};

export default connect(null, { addTypeUser })(withRouter(AddTypeUserForm));
