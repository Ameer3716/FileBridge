import {ImageResponse} from 'next/og';

export const alt='FileBridge — Upload here. Download anywhere.';
export const size={width:1200,height:630};
export const contentType='image/png';

export default function Image(){
 return new ImageResponse(
  <div style={{width:'100%',height:'100%',display:'flex',flexDirection:'column',justifyContent:'space-between',background:'#f6f8f2',padding:72,color:'#254c3b',fontFamily:'sans-serif'}}>
   <div style={{display:'flex',alignItems:'center',gap:20,fontSize:40,fontWeight:700}}>
    <svg width="72" height="72" viewBox="0 0 64 64"><rect width="64" height="64" rx="17" fill="#276b52"/><path d="M15 44V29a17 17 0 0 1 34 0v15M15 32h34M24 32v12m16-12v12" fill="none" stroke="#d4ebd6" strokeWidth="4" strokeLinecap="round"/><path d="m25 17 7-7 7 7m-7-6v15m-7 23 7 7 7-7m-7 6V40" fill="none" stroke="white" strokeWidth="4" strokeLinecap="round" strokeLinejoin="round"/></svg>
    FileBridge
   </div>
   <div style={{display:'flex',flexDirection:'column',gap:20}}>
    <div style={{fontSize:76,fontWeight:700,letterSpacing:-3}}>Upload here.</div>
    <div style={{fontSize:76,fontWeight:700,letterSpacing:-3,color:'#60936a'}}>Download anywhere.</div>
    <div style={{fontSize:27,color:'#637362',marginTop:12}}>Your files, between computers. No login needed.</div>
   </div>
   <div style={{display:'flex',fontSize:22,color:'#637362'}}>filebridge-red.vercel.app</div>
  </div>,size,
 );
}
