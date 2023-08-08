import React, { useCallback, useMemo, useState } from 'react';
import './App.css';
import { FormControl, Select, MenuItem, InputLabel, SelectChangeEvent, Button, TextField } from '@mui/material'
import  * as xlxs from 'xlsx';
import periodColumns from './constants/periodColumns';

import Papa from 'papaparse';


const allowedExtensions = ["csv"];

type SheetData = {
  label: string;
  firstRow: [string, string, string | undefined],
  data: {[key: string]: string}[],
}

type ParsedRow = {[key: string]: string};

const EXPORTED_FILE_TYPE = 'xls';

function App() {

  // This state will store the parsed data
  const [data, setData] = useState<ParsedRow[]>([]);
  // this state stores the columns saved from the csv
  const [fields, setFields] = useState<string[]>([]);
 
  const [error, setError] = useState<string>("");

  const [selectedTeacher, setSelectedTeacher] = useState<string>('');

  const [headerInput, setHeaderInput] = useState<string>('');

  const handleFileChange: React.ChangeEventHandler<HTMLInputElement> = (e) => {
      setError("");
      setSelectedTeacher('');
      setData([]);
      setFields([]);

      // Check if user has entered the file
      if (e?.target?.files?.length) {
          const inputFile = e.target.files[0];

          // Check the file extensions, if it not
          // included in the allowed extensions
          // we show the error
          const fileExtension = inputFile?.type.split("/")[1];
          if (!allowedExtensions.includes(fileExtension)) {
              setError("Please input a csv file");
              return;
          }

          try{
            Papa.parse(inputFile, {
              header: true,
              skipEmptyLines: true,
              complete: function(results){
                console.log(results.meta.fields);
                setFields(results.meta.fields || []);
                setData(results.data as ParsedRow[]);
              },
              error(error, file) {
                setError(error.message);
              },
            })
          } catch (e) {
            console.error(e);
            setError("An unknown error occured");
          }
      }
  };

  const teachers = useMemo(() => {
    const teacherSet = new Set<string>();
    for(let i = 0; i < data.length; i++){
      let dataPoint = data[i];
      for(let j = 0; j < periodColumns.length; j++){

        let columnData = dataPoint[periodColumns[j].columnLabel];
        if(columnData && !teacherSet.has(columnData)){
          teacherSet.add(columnData);
        }
      }
    }
    return Array.from(teacherSet);
  }, [data]);

  const fieldsForData = useMemo(() => {
    return fields.filter(field => !periodColumns.map(col => col.columnLabel).includes(field));
  }, [fields]);

  const handleChange = (event: SelectChangeEvent) => {
    setSelectedTeacher(event.target.value as string);
  };

  const handleInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setHeaderInput(event.target.value as string);
  };

  const generateTeacherData = useCallback(() => {

    const generatedSheets = new Set<SheetData>();

    const dataPerPeriod: Map<string, any[]> = new Map();

    periodColumns.forEach(period => {
      dataPerPeriod.set(period.columnLabel, []);
    });
    // for every record in the dataset
    for(let i = 0; i < data.length; i++){
      let dataPoint = data[i];
      // for every period that we are trying to parse these records into
      for(let j = 0; j < periodColumns.length; j++){
        const period = periodColumns[j];
        // if one of the periods is assigned to the selected teacher
        if(dataPoint[period.columnLabel] === selectedTeacher){
          
          // get the record minus any period data
          const dataPointLessPeriodData: ParsedRow = {}; 
          (new Map(Object.entries(dataPoint))).forEach((value, key) => {
            if(!periodColumns.some(periodData => periodData.columnLabel === key)){
              dataPointLessPeriodData[key] = value;
            }
          });

          // add it to any existing period data for the period for the teachers;
          dataPerPeriod.set(period.columnLabel, [
            ...(dataPerPeriod.get(period.columnLabel) || []), 
            dataPointLessPeriodData
          ])
        }
      }
    }

    // after going through the data create sheets
    [...dataPerPeriod.keys()].forEach(key => {
      let periodData = periodColumns.find(period => period.columnLabel === key);
      if(!periodData){
        setError(`An unknown error occurred generating sheet data, unable to find period data for column ${key}`);
      } else {
        generatedSheets.add({
          label: `${selectedTeacher.replace(/, .*/g, '')} ${periodData.period}`,
          firstRow: [selectedTeacher, periodData.period, headerInput],
          data: (dataPerPeriod.get(key) as any as {[key: string]: string}[]),
        })
      }
    });

    const workbook = xlxs.utils.book_new();

    generatedSheets.forEach(sheetData => {
      const worksheet = xlxs.utils.aoa_to_sheet<string | undefined>([
        sheetData.firstRow,
        [],
        fieldsForData,
        ...sheetData.data.map(data => fieldsForData.map(field => data[field]))
      ]);
      xlxs.utils.book_append_sheet(workbook, worksheet, sheetData.label.slice(0, 30));
    })

    xlxs.writeFile(workbook,  `${selectedTeacher}.${EXPORTED_FILE_TYPE}`);

  }, [data, selectedTeacher, headerInput, fieldsForData]);


  return (
    <div className="App">
      <Button
        component="label"
      >
        <input 
          type='file'
          accept=".csv"
          onChange={handleFileChange}
        />
      </Button>

      <FormControl fullWidth>
        <InputLabel id="teacher-select-label">Teacher</InputLabel>
        <Select
          labelId='teacher-select-label'
          id='teacher-select'
          value={selectedTeacher}
          label='Teacher'
          onChange={handleChange}
        >
          {teachers.map(teacher => {
            return <MenuItem key={teacher} value={teacher}>{teacher}</MenuItem>
          })};
        </Select>

        <TextField
          id='header-input'
          value={headerInput}
          label='Top Level Label'
          onChange={handleInputChange}
        />
      </FormControl>

      { error && (
        <p>
          {error}
        </p>
      )}

      <Button 
        variant='contained'
        onClick={generateTeacherData}
      > 
        Parse Data 
      </Button>
    </div>
  );
}

export default App;
