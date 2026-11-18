"""Unabhängige Nachrechnung des Flip-Rechners (zum Abgleich der Excel-Formeln)."""
def flip(K, wf, reno_qm, arv_qm, months, grest=0.05, notar=0.02, makler=0.0357, puffer=0.15, ltv=0.7, zins=0.045, lfd=300, vk=0.0357, vkfix=1500, tax=0.35, ziel=0.15, miete=0):
    nk=K*(grest+notar+makler); reno=wf*reno_qm*(1+puffer); basis=K+nk+reno
    i=ltv*zins*months/12; zinsen=basis*i; halte=zinsen+lfd*months
    arv=wf*arv_qm; vkost=arv*vk+vkfix
    gewinn=arv-vkost-basis-halte; ersp=miete*months; nach=gewinn*(1-tax) if gewinn>0 else gewinn
    ek=basis*(1-ltv)+halte
    maxK=(arv*(1-vk)-vkfix-reno*(1+i)-lfd*months+ersp-ziel*arv)/((1+grest+notar+makler)*(1+i))
    return dict(arv=round(arv),nk=round(nk),reno=round(reno),halte=round(halte),vk=round(vkost),gesamt=round(basis+halte+vkost),gewinn=round(gewinn),nach=round(nach),marge=round(gewinn/arv*100,1),ek=round(ek),maxK=round(maxK),abschlag=round((1-maxK/K)*100,1))
if __name__ == "__main__":
  deals={"Bungalow Damme-West":(250000,139.5,600,2700,8),"DHH Damme":(235000,113,450,2600,6),"EFH Lohne 1946":(195000,224,750,2200,10),"Lemförde Schimmel":(236000,149,900,2300,12),"ETW Damme 1993":(190000,83,300,2700,4)}
  for n,(K,wf,r,a,m) in deals.items(): print(n, flip(K,wf,r,a,m, lfd=250 if 'ETW' in n else 300))
