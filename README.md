# node-xlsx-handle
A handle for the xlsx package. The two main features are handleXlsx and convertXlsxToArray

**ConvertXlsxToArray** : 
 - Description : Takes the xlsx and return the matrix with the result
 - *Params:*
		Data: the file readed using FileReader
		(optional) Type: The format of the file (default base64)
 - *Return:*
		xlsx: The matrix
		header: The first row of the sheet
		ids: The index of the columns which are the ids of the doc

**handleXlsx** : 
 - Description : Takes the xlsx and return an array of objects with the data, according to the xlsx model
 - *Params:*
		Data: the file readed using FileReader
		(optional) Type: The format of the file (default base64)
 - *Return:*
		documents: The array of documents generated
