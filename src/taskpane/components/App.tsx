import * as React from "react";
import { Button, ButtonType, TextField, ITextFieldStyles, Slider, Checkbox } from "office-ui-fabric-react";
import { Stack, IStackTokens, Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from "office-ui-fabric-react";
import Header from "./Header";
import Progress from "./Progress";
/* global Button Header, HeroList, HeroListItem, Progress, Word */

const sliderValueFormat = (value: number) => `${value}px`;
const textfieldStyles: Partial<ITextFieldStyles> = {
  root: {width: 300}
};

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300, height: 50},
  
};

const options: IDropdownOption[] = [
  { key: 'fontsHeader', text: 'Fonts', itemType: DropdownMenuItemType.Header },
  { key: 'georgia', text: 'Georgia' },
  { key: 'elephant', text: 'Elephant' },
  { key: 'calibri', text: 'Calibri', },
  { key: 'rockwellExtraBold', text: 'Rockwell Extra Bold'},
  { key: 'stencil', text: 'Stencil'},
  { key: 'webdings', text: 'Webdings'}
];

const stackTokens: IStackTokens = { childrenGap: 10 };

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  headerText: string;
  fontSize: number;
  isBold: boolean;
  isItalic: boolean;
  isUnderline: boolean;
  selectedFont: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    
    this.state = {
      headerText: "",
      fontSize: 16,
      isBold: false,
      isItalic: false,
      isUnderline: false,
      selectedFont: "Calibri",
    }

  }

  giveMeHeader = async () => {
    return Word.run(async context => {
      
      const {fontSize, headerText, isBold, isItalic, isUnderline, selectedFont} = this.state;
      
      if(headerText === "") return;
      const paragraph = context.document.body.insertParagraph(headerText, Word.InsertLocation.start);
      
      paragraph.font.name = selectedFont;
      paragraph.font.color = "#000000";
      paragraph.font.size = fontSize;
      paragraph.font.bold = isBold;
      paragraph.font.italic = isItalic;
    
      isUnderline ? paragraph.font.underline = "Single" : paragraph.font.underline = "None"  

      await context.sync();
    });
  };

  changeTextValue = (ev) =>{
    this.setState({
      headerText: ev.target.value
    })
  }

  changeFontSize = (value: number) => {
    this.setState({
      fontSize: value
    })
  }

  changeFontValue = (ev, item) => {
    this.setState({
      selectedFont: item.text
    })
  }

  changeBold = (ev, checked) => {
    this.setState({
      isBold: checked
    })
  }

  changeItalic = (ev, checked) => {
    this.setState({
      isItalic: checked
    })
  }

  changeUnderline = (ev, checked) => {
    this.setState({
      isUnderline: checked
    })
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">

      <Stack tokens={stackTokens}>
        <Header message="Easy Plugin" />
        
        <TextField 
          label="Nadpis"
          styles={textfieldStyles} 
          value={this.state.headerText}
          placeholder="Text nadpisu..." 
          onChange={this.changeTextValue.bind(this)}
        />
        
        <Dropdown
          placeholder="Vybrat styl písma..."
          label="Vyberte si styl písma"
          options={options}
          styles={dropdownStyles}
          onChange={this.changeFontValue.bind(this)}
        />
        
        <Slider 
          label="Velikost Nadpisu" 
          min={16} 
          max={60} 
          step={1}
          onChange={this.changeFontSize}
          valueFormat={sliderValueFormat} 
          showValue
          snapToStep/>
          
        <Checkbox 
          label="tučné"
          onChange={this.changeBold.bind(this)}  
        />

        <Checkbox 
          label="kurzíva"
          onChange={this.changeItalic.bind(this)}  
        />

        <Checkbox 
          label="podtržení"
          onChange={this.changeUnderline.bind(this)}  
        />
          

        <Button
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          onClick={this.giveMeHeader}
        >Použít
        </Button>
      </Stack>

      </div>
    );
  }
}
