# VBA-Generate-Calendar
VBA macro to generate a calendar in Excel from a list of events.

Usage
-----

Import GenerateCalendar.bas into a new Excel VBA module.
Create 2 worksheets: Events and Calendar
Add your events list to the Events worksheet, using 1 header row. At minimum, you should have the following columns: Event Name, Event Date, Event Start Time, Event End Time, and Event Duration. [Here](https://docs.google.com/spreadsheets/d/13nmTGkXFL6PW17H03rXzOU6fHmeSnu-SOFysATPxFBQ/edit?usp=sharing) is a sample workbook to compare to.
Edit the `Const`s in the code as necessary to match the appropriate columns.
Run the Generate Calendar macro.

License
-------

MIT License

Copyright (c) 2016 Behind The Math

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.