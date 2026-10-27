import {list} from '@vercel/blob';
import {record} from '@/lib/files.mjs';
export const dynamic='force-dynamic';
export const runtime='nodejs';
export async function GET(){
 if(!process.env.BLOB_READ_WRITE_TOKEN)return Response.json({error:'One-time setup: connect a private Vercel Blob store in Vercel → Storage, then redeploy.'},{status:503});
 try{
  const files=[];let cursor:string|undefined;
  do{const page=await list({prefix:'files/',limit:1000,cursor});for(const blob of page.blobs){const f=record(blob);if(f)files.push(f)}cursor=page.hasMore?page.cursor:undefined;}while(cursor);
  files.sort((a,b)=>b.uploadedAt-a.uploadedAt||b.id.localeCompare(a.id));
  return Response.json({files},{headers:{'Cache-Control':'no-store'}});
 }catch{return Response.json({error:'Could not load files. Check the Blob storage connection and retry.'},{status:503})}
}
