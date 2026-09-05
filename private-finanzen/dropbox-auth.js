(() => {
  "use strict";
  const K={app:"financeDropboxAppKey",token:"financeDropboxAccessToken",exp:"financeDropboxTokenExpiresAt",ver:"financeDropboxCodeVerifier",state:"financeDropboxOAuthState"};
  let client=null;
  const redirectUri=()=>`${location.origin}${location.pathname}`;
  const appKey=()=>localStorage.getItem(K.app)||"";
  function saveAppKey(value){const key=String(value||"").trim();if(!key)throw new Error("Bitte zuerst den Dropbox App-Key eintragen.");localStorage.setItem(K.app,key);return key;}
  function token(){const t=localStorage.getItem(K.token), exp=Number(localStorage.getItem(K.exp)||0);if(!t||(exp&&Date.now()>exp-60000)){localStorage.removeItem(K.token);localStorage.removeItem(K.exp);client=null;return null;}return t;}
  async function getClient(){if(client)return client;if(!window.Dropbox)throw new Error("Die Dropbox-Bibliothek konnte nicht geladen werden.");const t=token();if(!t)return null;client=new Dropbox.Dropbox({accessToken:t});return client;}
  async function startLogin(keyValue){const key=saveAppKey(keyValue);const state=crypto.randomUUID();const auth=new Dropbox.DropboxAuth({clientId:key});const url=await auth.getAuthenticationUrl(redirectUri(),state,"code","online",null,"none",true);sessionStorage.setItem(K.ver,auth.getCodeVerifier());sessionStorage.setItem(K.state,state);location.assign(url);}
  async function completeLogin(){const q=new URLSearchParams(location.search),code=q.get("code");if(!code)return false;const key=appKey(),ver=sessionStorage.getItem(K.ver),expected=sessionStorage.getItem(K.state),returned=q.get("state");if(!key||!ver)throw new Error("Die Dropbox-Anmeldung konnte nicht abgeschlossen werden. Bitte erneut verbinden.");if(expected&&returned!==expected)throw new Error("Dropbox-Anmeldung abgebrochen: Sicherheitsprüfung fehlgeschlagen.");const auth=new Dropbox.DropboxAuth({clientId:key});auth.setCodeVerifier(ver);const response=await auth.getAccessTokenFromCode(redirectUri(),code);const result=response.result||response;if(!result.access_token)throw new Error("Dropbox hat kein Zugriffstoken zurückgegeben.");localStorage.setItem(K.token,result.access_token);localStorage.setItem(K.exp,String(Date.now()+Number(result.expires_in||14400)*1000));sessionStorage.removeItem(K.ver);sessionStorage.removeItem(K.state);history.replaceState({},document.title,redirectUri());client=new Dropbox.Dropbox({accessToken:result.access_token});return true;}
  function disconnect(){localStorage.removeItem(K.token);localStorage.removeItem(K.exp);client=null;}
  window.FinanceDropboxAuth={redirectUri,appKey,saveAppKey,getClient,startLogin,completeLogin,disconnect,isConnected:()=>Boolean(token())};
})();
