import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react";

interface IEmployeeDropdownProps {
  employees: IDropdownOption[];
  selectedEmployee: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
}

class EmployeeDropdown extends React.Component<IEmployeeDropdownProps, {}> {
  render() {
    const { employees, selectedEmployee, onChange } = this.props;
    const placeHolderText =
      employees.length === 0
        ? "All personnel have been evaluated. No more employees to evaluate."
        : "Choose an employee";

    return (
      <Dropdown
        placeHolder={placeHolderText}
        options={employees}
        onChanged={onChange}
        selectedKey={selectedEmployee}
      />
    );
  }
}

export default EmployeeDropdown;
