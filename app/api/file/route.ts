import {list,get,del} from '@vercel/blob';
import {UUID,record,downloadName,sameOrigin} from '@/lib/files.mjs';
export const dynamic='force-dynamic';
export const runtime='nodejs';
export const maxDuration=60;
const error=(message:string,status:number)=>Response.json({error:message},{status,headers:{'Cache-Control':'no-store'}});
async function find(request:Request){const id=new URL(request.url).searchParams.get('id')||'';if(!UUID.test(id))return null;const page=await list({prefix:`files/${id}/`,limit:2});return page.blobs.length===1?page.blobs[0]:null}
export async function GET(request:Request){
 try{
  const blob=await find(request);if(!blob)return error('File not found. Refresh the list.',404);
  const f=record(blob);if(!f)return error('Invalid file record.',404);
  const result=await get(blob.pathname,{access:'private',useCache:false});
  if(!result||result.statusCode!==200||!result.stream)return error('File not found.',404);
  const filename=downloadName(f.name,f.uploadedAt,f.id);
  const encoded=encodeURIComponent(filename).replace(/['()*]/g,c=>'%'+c.charCodeAt(0).toString(16));
  // Stream files; do not buffer large downloads inside the function.
  return new Response(result.stream,{headers:{'Content-Type':'application/octet-stream','Content-Disposition':`attachment; filename="download"; filename*=UTF-8''${encoded}`,'Cache-Control':'no-store','X-Content-Type-Options':'nosniff'}});
 }catch{return error('Download failed. Please retry.',500)}
}
export async function DELETE(request:Request){
 if(!sameOrigin(request))return error('Request not allowed.',403);
 try{const blob=await find(request);if(!blob)return error('File not found.',404);await del(blob.url);return Response.json({ok:true},{headers:{'Cache-Control':'no-store'}})}catch{return error('Could not delete file. Please retry.',500)}
}
