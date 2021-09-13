import * as React from "react";
import styles from "./HelloWorld.module.scss";
import { IHelloWorldProps } from "./IHelloWorldProps";
import { escape } from "@microsoft/sp-lodash-subset";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Checkbox, DatePicker, Dropdown, IBasePickerSuggestionsProps, IDropdownOption, IPersonaProps, NormalPeoplePicker, PrimaryButton, TextField } from "@fluentui/react";
import { sp } from "@pnp/sp";
import { IItemAddResult } from "@pnp/sp/items";
import { IListItemFormUpdateValue } from "@pnp/sp/lists";
import { PeoplePickerNormalExample } from "./CustomPeoplePicker";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

enum Operation {
  Add,
  Update
}

export const HelloWorld = (props: IHelloWorldProps) => {
  
  const urlParams = new URLSearchParams(window.location.search);
  const itemId = urlParams.get('ItemID');

  console.log("Props: ", props);
  console.log('TEST')
  const [title, setTitle] = React.useState<string>('');
  const [countryKey, setCountryKey] = React.useState<number>(7);
  const [date, setDate] = React.useState<Date>();
  const [selected, setSelected] = React.useState<boolean>(false);
  const [people, setPeople] = React.useState<any>(null);

  // SP Configuration
  React.useEffect(() => {
    sp.setup({
      sp: {
        baseUrl: props.context.pageContext.web.absoluteUrl,
      },
    });
  }, []);

  React.useEffect(() => {
    const getItem = async () => {
      if (itemId === null) { return; }
      const item: any = await sp.web.lists.getByTitle("ZentivaTest").items.getItemByStringId(itemId).get();
      console.log('Item with ID:', itemId, ' ',item);
      const countryKey = options.filter(s => s.text === item.Country)[0].key;
      const user = await sp.web.getUserById(item.ResponsibleId).get();
      console.log('User', user);
      setTitle(item.Title);
      setCountryKey(countryKey as number);
      setDate(new Date(item.Birthday));
      setSelected(item.Selected);
      setPeople(user);
    }
    getItem();
  }, [])

  // Get list items
  React.useEffect(() => {
    const getItems = async () => {
      const items: any[] = await sp.web.lists.getByTitle("ZentivaTest").items.get();
      console.log(`Items: ${items}`);
      // update
    }
    getItems();
  }, []);

  const options : IDropdownOption[] = [
    {key: 1, text: 'USA'},
    {key: 2, text: 'Canada'},
    {key: 3, text: 'Mexico'},
    {key: 4, text: 'Czech Republic'},
    {key: 5, text: 'Slovakia'},
    {key: 6, text: 'Germany'},
    {key: 7, text: ''},
  ]

  const onDateChange = (date: Date | null | undefined) => {
    setDate(date);
  }
  const onTextfieldChange = (event: any, newValue?: string) => {
    setTitle(newValue);
  }

  const onDropdownChange = (event:any, option?: IDropdownOption, index?: number) => {
    setCountryKey(option.key as number);
  }

  const onCheckboxChange = (ev?: any, checked?: boolean) => {
    setSelected(checked);
  }

  const onPeopleChange = (items: any[]) => {
    console.log('onPeopleChange' , items);
    setPeople(items[0]);
  }

  const getLoginName = (user) => {
    if (user === null || user === undefined) {
      return '';
    }
    return user.loginName ? user.loginName : user.LoginName;
  }

  const getEmail = (user) => {
    if (user === null || user === undefined) {
      return '';
    }
    return user.Email ? user.Email : user.secondaryText;
  }

  const onButtonClick = async (operation : Operation) => {
    const dateInLocale = date.toLocaleDateString().split('/');
    const formattedDate = `${dateInLocale[1]}.${dateInLocale[0]}.${dateInLocale[2]}`;
    const selectedCountryText = options.filter(s => s.key === countryKey)[0].text;
    const loginName = getLoginName(people); 
    const formValues: IListItemFormUpdateValue[] = [
      {FieldName: 'Title', FieldValue: title},
      {FieldName: 'Country', FieldValue: selectedCountryText},
      {FieldName: 'Birthday', FieldValue: formattedDate},
      {FieldName: 'Selected', FieldValue: selected ? "true" : "false"},
      {FieldName: 'Responsible', FieldValue: JSON.stringify([{ Key: loginName }])},
    ]
    let response : IListItemFormUpdateValue[] = [];
    switch(operation) {
      case Operation.Add:
        response = await sp.web.lists.getByTitle("ZentivaTest").addValidateUpdateItemUsingPath(formValues, '');
        break;
      case Operation.Update:
        response = await sp.web.lists.getByTitle("ZentivaTest").items.getItemByStringId(itemId).validateUpdateListItem(formValues);
        break;
    }
    console.log("Response", response);
  }
// Gulp clean; Gulp bundle --max_old_space_size=8192 --ship; Gulp package-solution --ship
  React.useEffect(() => {
    console.log("State: ", people, title , countryKey , date, selected); 
  }, [title, countryKey, date, selected, people])
  console.log('loginName', getLoginName(people));
  return (
    <div className={styles.helloWorld}>
      <div className={styles.container}>
        <div className={styles.row}>
          <TextField label="Title" onChange={onTextfieldChange} value={title}/>
          <Dropdown label="Country" onChange={onDropdownChange} options={options} selectedKey={countryKey} />
          <DatePicker label={'Birthday'} onSelectDate={onDateChange} value={date}/>
          <Checkbox label="Selected" className={styles.customCheckbox} checked={selected} onChange={onCheckboxChange} />
          <PeoplePicker
            context={props.context}
            titleText="Responsible"
            personSelectionLimit={props.slider}
            groupName={''} // Leave this blank in case you want to filter from all users
            showtooltip={true}
            onChange={onPeopleChange}
            showHiddenInUI={false}
            defaultSelectedUsers={[getEmail(people)]}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000} 
          />
          <PrimaryButton text="Send data" style={{marginTop: '15px'}} onClick={() => onButtonClick(Operation.Add)} allowDisabledFocus />
          <PrimaryButton text="Edit data" style={{marginTop: '15px'}} onClick={() => onButtonClick(Operation.Update)} allowDisabledFocus />
        </div>
      </div>
    </div>
  );
};
