import {mkdir,cp,writeFile} from 'node:fs/promises';
await mkdir('dist',{recursive:true});
for(const path of ['index.html','styles.css','src','assets','manifest.xml']) await cp(path,`dist/${path}`,{recursive:true});
await writeFile('dist/.nojekyll','');
console.log('Built static app in dist/');
