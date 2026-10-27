import {handleUpload,type HandleUploadBody} from '@vercel/blob/client';
import {parsePath,MAX_BYTES,sameOrigin} from '@/lib/files.mjs';
export const runtime='nodejs';
export async function POST(request:Request){
 if(!process.env.BLOB_READ_WRITE_TOKEN)return Response.json({error:'One-time setup: connect a private Vercel Blob store, then redeploy.'},{status:503});
 if(!sameOrigin(request))return Response.json({error:'Request not allowed.'},{status:403});
 try{
  const body=await request.json() as HandleUploadBody;
  const result=await handleUpload({body,request,
   onBeforeGenerateToken:async(pathname)=>{
    // This is deliberately a no-account file shelf, as requested. A visitor
    // who can open the website can upload, list, download, and delete files.
    if(!parsePath(pathname))throw Error('Invalid file name or upload path.');
    return {maximumSizeInBytes:MAX_BYTES,addRandomSuffix:false,allowOverwrite:false,validUntil:Date.now()+15*60000};
   },
   onUploadCompleted:async()=>{ /* Blob itself is the durable file index. */ },
  });
  return Response.json(result,{headers:{'Cache-Control':'no-store'}});
 }catch{return Response.json({error:'Upload could not start. Check your storage connection or try again.'},{status:400})}
}
