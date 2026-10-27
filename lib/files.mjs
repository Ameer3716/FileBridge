export const MAX_BYTES=25*1024*1024;
export const UUID=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export function parsePath(pathname){
 const parts=pathname.split('/');
 if(parts.length!==5||parts[0]!=='files'||!UUID.test(parts[1])||!/^\d{1,16}$/.test(parts[2])||!/^[a-f0-9]{64}$/.test(parts[3]))return null;
 try{const name=decodeURIComponent(parts[4]);if(!name||name.length>180||/[\x00-\x1f\x7f\\/]/.test(name))return null;return {id:parts[1],modifiedAt:Number(parts[2])||null,sha256:parts[3],name}}catch{return null}
}
export function makePath(id,modifiedAt,hash,name){
 const safe=name.replace(/[\x00-\x1f\x7f\\/]/g,'_').slice(0,180).trim()||'untitled';
 return `files/${id}/${Math.max(0,Math.trunc(modifiedAt)||0)}/${hash}/${encodeURIComponent(safe)}`;
}
export function record(blob){const data=parsePath(blob.pathname);return data?{...data,size:blob.size,uploadedAt:new Date(blob.uploadedAt).getTime()}:null}
export function downloadName(name){return name}
export function sameOrigin(request){const origin=request.headers.get('origin');return !origin||origin===new URL(request.url).origin}
