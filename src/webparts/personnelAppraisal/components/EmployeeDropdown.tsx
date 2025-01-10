import * as React from "react";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react";

export interface IEmployeeDropdownProps {
  employees: IDropdownOption[];
  selectedEmployee: string | number | undefined;
  onChange: (option?: IDropdownOption) => void;
}

class EmployeeDropdown extends React.Component<IEmployeeDropdownProps, {}> {
  render() {
    const { employees, selectedEmployee, onChange } = this.props;
    const placeHolderText =
      employees.length === 0
        ? "همه افراد لیست شما برای این دوره ارزیابی شده اند."
        : "ارزیابی شونده را انتخاب کنید";
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
