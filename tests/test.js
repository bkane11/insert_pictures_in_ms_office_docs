// test.js
'use strict';

const addImages = require('../index.js')
, fs = require('fs')
, images = [
  './tests/images/image_1.jpg'
  , './tests/images/image_2.jpg'
  , './tests/images/image_does_not_exist.jpg'
  ]
, xlsx = './tests/outputs/test.xlsx'
, docx = './tests/outputs/test.docx'
;

Promise.all(
  [xlsx, docx].map(file=>{
    if(fs.existsSync(file))
      fs.unlinkSync(file)

    return new Promise((resolve, reject)=>{
      fs.createReadStream(file.replace('outputs', 'inputs'))
        .on('error', err=>reject(err))
        .pipe(
          fs.createWriteStream(file)
            .on('close', evt=>resolve(file))
            .on('error', err=>reject(err))
        )
    })
  })
)
.then(doTest)
.catch(err=>{throw err})

function doTest(){
  addImages(xlsx, images)
    .then(filepath=>console.log('success!  images added to:', filepath))
    .catch(err=>console.log('error:', err))
  
  addImages(docx, images)
    .then(filepath=>console.log('success!  images added to:', filepath))
    .catch(err=>console.log('error:', err))

}