import type {Metadata} from 'next';
import './globals.css';
export const metadata:Metadata={title:'FileBridge — Files between computers',description:'Upload here. Download on your other computer. No login needed.',robots:{index:false,follow:false}};
export default function Layout({children}:{children:React.ReactNode}){return <html lang="en"><body>{children}</body></html>}
