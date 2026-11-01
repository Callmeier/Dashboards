(() => {
  "use strict";
  const K={
    app:"financeDropboxAppKey",
    token:"financeDropboxAccessToken",
    exp:"financeDropboxTokenExpiresAt",
    refresh:"financeDropboxRefreshToken",
    ver:"financeDropboxCodeVerifier",
    state:"financeDropboxOAuthState"
  };
  let client=null;
  const redirectUri=()=>`${location.origin}${location.pathname}`;
  const DEFAULT_APP_KEY="vqm3q95w02p32hl";
  const appKey=()=>localStorage.getItem(K.app)||DEFAULT_APP_KEY;
  const refreshToken=()=>localStorage.getItem(K.refresh)||"";
  function saveAppKey(value){
    const key=String(value||"").trim();
    if(!key)throw new Error("Bitte zuerst den Dropbox App-Key eintragen.");
    localStorage.setItem(K.app,key);
    return key;
  }
  function storedAccess(){
    const t=localStorage.getItem(K.token)||"";
    const exp=Number(localStorage.getItem(K.exp)||0);
    return {token:t,expiresAt:exp};
  }
  function validAccess(){
    const a=storedAccess();
    return Boolean(a.token&&(!a.expiresAt||Date.now()<a.expiresAt-60000));
  }
  function persistAuth(auth){
    const t=auth.getAccessToken?.();
    const exp=auth.getAccessTokenExpiresAt?.();
    const ref=auth.getRefreshToken?.();
    if(t)localStorage.setItem(K.token,t);
    if(exp instanceof Date&&!Number.isNaN(exp.getTime()))localStorage.setItem(K.exp,String(exp.getTime()));
    if(ref)localStorage.setItem(K.refresh,ref);
  }
  async function getClient(){
    if(client)return client;
    if(!window.Dropbox)throw new Error("Die Dropbox-Bibliothek konnte nicht geladen werden.");
    const key=appKey();
    if(!key)return null;
    const a=storedAccess(),ref=refreshToken();
    if(!ref&&!validAccess())return null;
    const auth=new Dropbox.DropboxAuth({
      clientId:key,
      accessToken:a.token||undefined,
      accessTokenExpiresAt:a.expiresAt?new Date(a.expiresAt):undefined,
      refreshToken:ref||undefined
    });
    try{
      if(ref){
        if(typeof auth.checkAndRefreshAccessToken==="function")await auth.checkAndRefreshAccessToken();
        else if(!validAccess()&&typeof auth.refreshAccessToken==="function")await auth.refreshAccessToken();
        persistAuth(auth);
      }
    }catch(e){
      console.error("Dropbox Token-Erneuerung",e);
      throw new Error("Die Dropbox-Verbindung konnte nicht automatisch erneuert werden. Bitte Dropbox einmal neu verbinden.");
    }
    if(!auth.getAccessToken?.())return null;
    client=new Dropbox.Dropbox({auth});
    return client;
  }
  async function startLogin(keyValue){
    const key=saveAppKey(keyValue);
    const state=crypto.randomUUID();
    const auth=new Dropbox.DropboxAuth({clientId:key});
    const url=await auth.getAuthenticationUrl(redirectUri(),state,"code","offline",null,"none",true);
    sessionStorage.setItem(K.ver,auth.getCodeVerifier());
    sessionStorage.setItem(K.state,state);
    location.assign(url);
  }
  async function completeLogin(){
    const q=new URLSearchParams(location.search),code=q.get("code");
    if(!code)return false;
    const key=appKey(),ver=sessionStorage.getItem(K.ver),expected=sessionStorage.getItem(K.state),returned=q.get("state");
    if(!key||!ver)throw new Error("Die Dropbox-Anmeldung konnte nicht abgeschlossen werden. Bitte erneut verbinden.");
    if(expected&&returned!==expected)throw new Error("Dropbox-Anmeldung abgebrochen: Sicherheitsprüfung fehlgeschlagen.");
    const auth=new Dropbox.DropboxAuth({clientId:key});
    auth.setCodeVerifier(ver);
    const response=await auth.getAccessTokenFromCode(redirectUri(),code);
    const result=response.result||response;
    if(!result.access_token)throw new Error("Dropbox hat kein Zugriffstoken zurückgegeben.");
    auth.setAccessToken(result.access_token);
    auth.setAccessTokenExpiresAt(new Date(Date.now()+Number(result.expires_in||14400)*1000));
    if(result.refresh_token)auth.setRefreshToken(result.refresh_token);
    persistAuth(auth);
    sessionStorage.removeItem(K.ver);
    sessionStorage.removeItem(K.state);
    history.replaceState({},document.title,redirectUri());
    client=new Dropbox.Dropbox({auth});
    return true;
  }
  function disconnect(){
    localStorage.removeItem(K.token);
    localStorage.removeItem(K.exp);
    localStorage.removeItem(K.refresh);
    client=null;
  }
  window.FinanceDropboxAuth={
    redirectUri,appKey,saveAppKey,getClient,startLogin,completeLogin,disconnect,
    isConnected:()=>Boolean(refreshToken()||validAccess())
  };
})();