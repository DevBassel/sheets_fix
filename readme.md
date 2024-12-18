# first step

### install [nodejs](https://nodejs.org/en/download/prebuilt-installer)

# step 2

### clone project from repo

```sh
git clone https://github.com/DevBassel/sheets_fix.git
```

#

# step 3 **IMPORTANT**

### copy rows from main file to new witout styles

# step 4

### install Dependencies with:

```sh
> npm i -g pnpm
> pnpm i
```

# step 5

open `app.ts` and look at

```sh
handelStuSheets("path/to/file.xlsx", "export file name");
```

# step 6

### Run App

```sh
> pnpm start
```

### data will be exported in `./data`

#

## If You Need Convert Excel To CSV

```js
convertExcelToCSV("path/to/file.xlsx", "export file name");
```

### and run app
