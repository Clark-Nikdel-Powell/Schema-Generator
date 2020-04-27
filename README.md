# Schema-Generator
Script for generating Schema in JSON-LD syntax from a Google Sheet.
Example: https://docs.google.com/spreadsheets/d/1jIAQ-Pmhzp-fvIUyyuoeuZD4gf0EcXSqoFbM3XIJPZI/edit#gid=0

This script, when added through the Script Editor in a Google Spreadsheet, will watch onEdit and onOpen to add two features to a Schema template.

## Advanced Data Validation

The onEdit function watches the Schema Type column. When changes are made, it finds the selected Type in the "Types" sheet, then uses the Properties column to fill in the available properties in the Schema Property column.

## Schema.org JSON-LD Generation

The onOpen function adds a button to a new Menu, called "SCHEMA." Clicking the "Generate Schema" button triggers the callback function, which grabs all of the data in the spreadsheet, then builds Schema.org JSON-LD markup based on the data.

### Schema Generation Example

#### Data

| Schema ID | Schema Type | Schema Property | Schema Value |
|---|---|---|---|
| #bakercollege | CollegeOrUniversity | name | Baker College  |
|   |   | logo | https://www.baker.edu/assets/images/logo/new-logo.svg |
|   |   | url  | https://www.baker.edu |
|   |   | telephone | (800) 964-4299 |
|   |   | sameAs | https://www.facebook.com/bakercollege,https://twitter.com/bakercollege,https://www.instagram.com/bakercollegeofficial/,https://www.youtube.com/user/baker,https://www.linkedin.com/school/baker-college/ |
|   |   | address | #address |
|   |   | alumni | #alumni |
| #address | PostalAddress | streetAddress | 1020 South Washington Street |
|   |   | addressLocality | Flint |
|   |   | addressRegion | Michigan |
|   |   | postalCode | 48867 |
| #alumni | Person | givenName | Eldon |
|   |   | familyName | Baker |

#### JSON-LD

```html
<script type="application/ld+json">
{
	"@context": "http://schema.org",
	"@id": "#bakercollege",
	"@type": "CollegeOrUniversity",
	"name": "Baker College",
	"logo": "https://www.baker.edu/assets/images/logo/new-logo.svg",
	"url": "https://www.baker.edu",
	"telephone": "(800) 964-4299",
	"sameAs": [
		"https://www.facebook.com/bakercollege",
		"https://twitter.com/bakercollege",
		"https://www.instagram.com/bakercollegeofficial/",
		"https://www.youtube.com/user/baker",
		"https://www.linkedin.com/school/baker-college/"
	],
	"address": {
		"@id": "#address"
	},
	"alumni": {
		"@id": "#alumni"
	}
}
</script>
<script type="application/ld+json">
{
	"@context": "http://schema.org",
	"@id": "#address",
	"@type": "PostalAddress",
	"streetAddress": "1020 South Washington Street",
	"addressLocality": "Flint",
	"addressRegion": "Michigan",
	"postalCode": "48867"
}
</script>
<script type="application/ld+json">
{
	"@context": "http://schema.org",
	"@id": "#alumni",
	"@type": "Person",
	"givenName": "Eldon",
	"familyName": "Baker"
}
</script>

```

#### Structured Data Testing Tool Results

## Feedback / Issues

Want to say hi? Email the author at josh@cnpagency.com. Got some comments, questions and/or concerns? Hit us up in the [Issues tab](https://github.com/Clark-Nikdel-Powell/Schema-Generator/issues).
