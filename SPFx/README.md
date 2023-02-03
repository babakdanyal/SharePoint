# SharePoint Framework

## Summary

This directory will contain SharePoint Framework code samples

## Used SharePoint Framework Version

![1.11.0](https://img.shields.io/badge/version-1.11-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Solution

Solution|Author(s)
--------|---------
spfx | Muhammad Babak Danyal Awan (babak.danyal@purple.telstra.com)

## Version history

Version|Date|Comments
-------|----|--------
1.0|03-February-2023|release

## Minimal Path to Awesome

- Clone this repository
- `npm install`
- update the serve.json file in the config folder with your sharepoint site collection
- `gulp serve`
```

### Build and deploy

  - `npm i`
  - `gulp clean`
  - `gulp bundle --ship`
  - `gulp build --ship`
  - `gulp package-solution --ship`
  - Deploy the app package (`sharepoint/solution/<package name>.spkg`) to your tenant AppCatalog