import type {Metadata} from 'next';
import './globals.css';
const title='FileBridge — Files between computers';
const description='Upload on one computer. Download on another. Share files with one simple link, no login needed.';
export const metadata:Metadata={
 metadataBase:new URL('https://filebridge-red.vercel.app'),
 title,description,applicationName:'FileBridge',
 alternates:{canonical:'/'},
 icons:{icon:[{url:'/favicon.svg',type:'image/svg+xml'}]},
 openGraph:{type:'website',locale:'en_US',url:'/',siteName:'FileBridge',title,description},
 twitter:{card:'summary_large_image',title,description,images:[{url:'/opengraph-image',alt:'FileBridge — Upload here. Download anywhere.'}]},
 robots:{index:false,follow:false},
};
export default function Layout({children}:{children:React.ReactNode}){return <html lang="en"><body>{children}</body></html>}
