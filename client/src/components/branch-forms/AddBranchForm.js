import React, { Fragment, useState, useEffect } from "react";
import { Link, withRouter } from "react-router-dom";
import PropTypes from "prop-types";
import { connect } from "react-redux";
import { addBranch, getAllBranchs } from "../../actions/branch";

const AddBranchForm = ({ addBranch, editing, currentBranch, history }) => {
  const initialFormState = { id: null, branchName: "", branchAddress: "" };

  const [formData, setFormData] = useState(initialFormState);

  const onChange = (e) =>
    setFormData({ ...formData, [e.target.name]: e.target.value });

  const clearForm = () => {
    setFormData({ ...formData, branchName: "", branchAddress: "" });
  };

  return (
    <Fragment>
      {/* <form
        class="form"
        onSubmit={(e) => {
          e.preventDefault();
          addBranch(formData, history);
          clearForm();
        }}
      >
        <div className="form-group add-flex" >
          <div class="col-md-4">
            <input
              type="text"
              placeholder="* Tên chi nhánh"
              name="branchName"
              value={formData.branchName}
              onChange={(e) => onChange(e)}
              style={{ fontSize: '14px' }}
              required
            />

          </div>
          <div class="col-md-4" >
            <input
              type="text"
              placeholder="* Địa chỉ"
              name="branchAddress"
              value={formData.branchAddress}
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
            addBranch(formData, history);
            clearForm();
          }}
        >
          <div className="form-body">
            <div className="row p-t-5 m-l-5">
              <div className="col-md-addJob-4">
                <div className="form-group has-success">
                  <input
                    type="text"
                    placeholder="* Tên chi nhánh"
                    name="branchName"
                    className="form-control"
                    value={formData.branchName}
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
                    placeholder="* Địa chỉ"
                    name="branchAddress"
                    className="form-control"
                    value={formData.branchAddress}
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

AddBranchForm.propTypes = {
  addBranch: PropTypes.func.isRequired,
  getAllBranchs: PropTypes.func.isRequired,
};

export default connect(null, { addBranch, getAllBranchs })(
  withRouter(AddBranchForm)
);
