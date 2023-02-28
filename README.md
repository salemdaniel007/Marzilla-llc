## To run this program

Install dependencies
```
npm install 
```

run the code
```
npx ts-node merged-cells.ts path/to/your/file.xlsx
```

Expected result for sample file
```
{
  Sheet1: [
    {
      s: { r: 0, c: 4 },
      e: { r: 0, c: 5 },
    },
  ],
};
```