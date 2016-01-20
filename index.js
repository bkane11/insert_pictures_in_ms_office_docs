'use strict';
const PythonShell = require('python-shell');


function addImages(xlsx, images){
  console.log('images', images, typeof images)
  let options = {
    mode: 'text',
    // pythonPath: 'path/to/python',
    pythonOptions: ['-u'],
    scriptPath: './python',
    args: [xlsx, '--images', images]
  };

  return new Promise((resolve, reject)=>{
    PythonShell.run('add_image_module.py', options, function (err, results) {
      if (err) 
        return reject(err);

      for(let result of results){
        try{
          // console.log(result, typeof result)
          result = JSON.parse(result.replace(/'|`/g, '"'));
          // console.log(result)
          if(result.success)
            return resolve(result.success)
        }catch(err){ 
          // console.log('error', err)
        }
      }

      return reject('no success item found')
      // console.log('results: %j', results);
    });
  })
}

module.exports = addImages;
