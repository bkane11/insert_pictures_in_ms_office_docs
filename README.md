# insert_pictures_in_ms_office_docs
tools for inserting pictures in MS Office Docs (xlsx, docx) using python and nodejs

```npm install insert_pictures_in_ms_office_docs```

returns a Promise that whe successful resolves with the filepath of the updated file that images were added to

install dependencies  ```npm install```

test with ```npm test```

use like: 
```javascript
const addImages = require('insert_pictures_in_ms_office_docs')
, images = [
  './tests/images/image_1.jpg'
  , './tests/images/image_2.jpg'
  , './tests/images/image_does_not_exist.jpg'
  ]
, xlsx = './tests/outputs/test.xlsx'
, docx = './tests/outputs/test.docx'
;

addImages(xlsx, images)
  .then(filepath=>console.log('success!  images added to:', filepath))
  .catch(err=>console.log('error:', err))
  
addImages(docx, images)
  .then(filepath=>console.log('success!  images added to:', filepath))
  .catch(err=>console.log('error:', err))
```
