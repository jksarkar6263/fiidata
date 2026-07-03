cat > /mnt/user-data/outputs/stock-widget/sector-widget.js << 'EOF'
/* ══════════════════════════════════════════════════════════
   sector-widget.js
   Host at: jksarkar6263/fiidata@main/marketMovers/sector-widget.js
   ══════════════════════════════════════════════════════════ */
function initSectorWidget(){
  var BASE="https://all-in-one.stockmarketsinindia.workers.dev/api/marketMovers";
  var ENDPOINTS={
    NSE:BASE+"/nse/SectoralPerformance",
    BSE:BASE+"/bse/SectoralPerformance"
  };
  var widget=document.getElementById("sec-widget");
  var cache={};
  var currentExchange="NSE";
  var currentView="table";   /* default = table */
  var sortCol="MCAP";
  var sortAsc=false;
  /* ── colour for heatmap tiles ── */
  function tileColor(p){
    p=parseFloat(p)||0;
    if(p> 4)return"#1b5e20";if(p> 2)return"#2e7d32";if(p> 0.5)return"#43a047";
    if(p>-0.5)return"#757575";if(p>-2)return"#e53935";if(p>-4)return"#c62828";
    return"#7f0000";
  }
  function fmt(v,d){if(v==null||v==="")return"—";var n=Number(v);return isNaN(n)?String(v):n.toFixed(d!=null?d:2);}
  function sCls(v){var n=Number(v);return isNaN(n)||n===0?"":n>0?"s-pos":"s-neg";}
  function sFmt(v,d){var n=Number(v);if(isNaN(n))return"—";return(n>0?"&#9650; ":n<0?"&#9660; ":"")+Math.abs(n).toFixed(d!=null?d:2);}
  function arrow(col){if(sortCol!==col)return" &#8645;";return sortAsc?" &#9650;":" &#9660;";}
  function sortedData(data){
    return data.slice().sort(function(a,b){
      if(sortCol==="Name")return sortAsc
        ?String(a.Name).localeCompare(String(b.Name))
        :String(b.Name).localeCompare(String(a.Name));
      return sortAsc
        ?(Number(a[sortCol])||0)-(Number(b[sortCol])||0)
        :(Number(b[sortCol])||0)-(Number(a[sortCol])||0);
    });
  }
  /* ── A/D bar cell (float-based, Edge-safe) ── */
  function adBar(up,nc,dn){
    var tot=(up+nc+dn)||1;
    var upW=Math.round(up/tot*100);
    var ncW=Math.round(nc/tot*100);
    var dnW=100-upW-ncW;
    return '<td class="ad-cell">'
      +'<div class="ad-nums">'
        +'<span class="an-up">&#9650;'+up+'</span>'
        +'<span class="an-nc">&#8212;'+nc+'</span>'
        +'<span class="an-dn">&#9660;'+dn+'</span>'
      +'</div>'
      +'<div class="ad-bar" style="width:100%;height:10px;border-radius:4px;overflow:hidden;background:#e8e8e8">'
        +'<div style="width:'+upW+'%;height:100%;background:#4caf50;float:left"></div>'
        +'<div style="width:'+ncW+'%;height:100%;background:#ffa726;float:left"></div>'
        +'<div style="width:'+dnW+'%;height:100%;background:#ef5350;float:left"></div>'
      +'</div>'
      +'</td>';
  }
  /* ── top control bar (exchange tabs + view toggle) ── */
  function controlBar(exchange,view,tUp,tDn,tNc){
    return '<div class="sec-ctrl">'
      /* left: exchange tabs */
      +'<div class="sec-exch">'
        +'<button class="sec-tab'+(exchange==="NSE"?" active":"")
          +'" onclick="secSwitch(\'NSE\')">NSE</button>'
        +'<button class="sec-tab'+(exchange==="BSE"?" active":"")
          +'" onclick="secSwitch(\'BSE\')">BSE</button>'
      +'</div>'
      /* center: title + summary */
      +'<div class="sec-mid">'
        +'<span class="sec-exch-label">'+exchange+' Sectoral Performance</span>'
        +'<span class="sec-adv-dec">'
          +'<span class="s-up">&#9650; '+tUp+' Adv</span>'
          +'<span class="s-dn">&#9660; '+tDn+' Dec</span>'
          +'<span class="s-nc">&#8212; '+tNc+' Unch</span>'
        +'</span>'
      +'</div>'
      /* right: view toggle */
      +'<div class="sec-view-btns">'
        +'<button class="sec-view-btn'+(view==="card"?" active":"")
          +'" onclick="secView(\'card\')">&#9638; Card</button>'
        +'<button class="sec-view-btn'+(view==="table"?" active":"")
          +'" onclick="secView(\'table\')">&#9776; Table</button>'
      +'</div>'
    +'</div>';
  }
  /* ── heatmap ── */
  function buildHeatmap(data){
    var sorted=data.slice().sort(function(a,b){
      return(Number(b.MCAP_PerChange)||0)-(Number(a.MCAP_PerChange)||0);
    });
    return '<div class="sec-heatmap">'
      +sorted.map(function(r){
        var pct=parseFloat(r.MCAP_PerChange)||0;
        var col=tileColor(pct);
        var sign=pct>0?"&#9650; ":pct<0?"&#9660; ":"";
        return '<div class="sec-tile" style="background:'+col+'">'
          +'<div class="sec-tile-name">'+r.Name+'</div>'
          +'<div class="sec-tile-chg">'+sign+Math.abs(pct).toFixed(2)+'%</div>'
          +'<div class="sec-tile-sub">&#916; &#8377;'+fmt(r.MCAP_CHANGE)+'Cr</div>'
          +'<div class="sec-tile-ad">&#9650;'+(r.UP||0)+' &#9660;'+(r.down||0)+' &#8212;'+(r.noChg||0)+'</div>'
          +'</div>';
      }).join("")
    +'</div>';
  }
  /* ── table ── */
  function buildTable(data){
    var rows=sortedData(data).map(function(r){
      return '<tr>'
        +'<td>'+r.Name+'</td>'
        +adBar(r.UP||0,r.noChg||0,r.down||0)
        +'<td class="'+sCls(r.ADRatio-1)+'">'+fmt(r.ADRatio)+'</td>'
        +'<td>&#8377;'+Number(r.MCAP||0).toLocaleString("en-IN",{maximumFractionDigits:0})+'</td>'
        +'<td class="'+sCls(r.MCAP_CHANGE)+'">&#8377;'+fmt(r.MCAP_CHANGE)+'Cr</td>'
        +'<td class="'+sCls(r.MCAP_PerChange)+'">'+sFmt(r.MCAP_PerChange)+'%</td>'
     +'</tr>';
    }).join("");
    return '<div class="sec-scroll">'
      +'<table class="sec-table"><thead><tr>'
        +'<th class="sortable" style="text-align:center" onclick="secSort(\'Name\')">Sector'+arrow("Name")+'</th>'
        +'<th style="text-align:left">Adv / Unch / Dec</th>'
        +'<th>A/D Ratio</th>'
        +'<th class="sortable" onclick="secSort(\'MCAP\')">MCap (Cr)'+arrow("MCAP")+'</th>'
        +'<th class="sortable" onclick="secSort(\'MCAP_CHANGE\')">MCap Chg (Cr)'+arrow("MCAP_CHANGE")+'</th>'
        +'<th class="sortable" onclick="secSort(\'MCAP_PerChange\')">MCap Chg%'+arrow("MCAP_PerChange")+'</th>'
        +'</tr></thead><tbody>'+rows+'</tbody></table>'
    +'</div>';
  }
  /* ── main render ── */
  function render(exchange,data){
    var tUp=0,tDn=0,tNc=0;
    data.forEach(function(r){tUp+=r.UP||0;tDn+=r.down||0;tNc+=r.noChg||0;});
    var content=controlBar(exchange,currentView,tUp,tDn,tNc);
    content+=(currentView==="card")?buildHeatmap(data):buildTable(data);
    widget.innerHTML=content;
  }
  /* ── fetch + cache ── */
  async function load(exchange){
    if(!cache[exchange]){
      widget.innerHTML=
        '<div class="sec-ctrl">'
          +'<div class="sec-exch">'
            +'<button class="sec-tab'+(exchange==="NSE"?" active":"")
              +'" onclick="secSwitch(\'NSE\')">NSE</button>'
            +'<button class="sec-tab'+(exchange==="BSE"?" active":"")
              +'" onclick="secSwitch(\'BSE\')">BSE</button>'
          +'</div>'
          +'<div class="sec-mid"></div>'
          +'<div class="sec-view-btns">'
            +'<button class="sec-view-btn'+(currentView==="card"?" active":"")
              +'" onclick="secView(\'card\')">&#9638; Card</button>'
            +'<button class="sec-view-btn'+(currentView==="table"?" active":"")
              +'" onclick="secView(\'table\')">&#9776; Table</button>'
          +'</div>'
        +'</div>'
        +'<div class="sec-loading"><span class="sec-spinner"></span> Loading '+exchange+' sector data...</div>';
      try{
        var res=await fetch(ENDPOINTS[exchange]);
        var json=await res.json();
        if(!json.success||!json.data||!json.data.length)throw new Error("No data");
        cache[exchange]=json.data;
      }catch(e){
        widget.innerHTML='<p class="sec-err">Could not load '+exchange+' sector data: '+e.message+'</p>';
        return;
      }
    }
    render(exchange,cache[exchange]);
  }
  /* ── global handlers ── */
  window.secSwitch=function(ex){currentExchange=ex;load(ex);};
  window.secView  =function(v){currentView=v;if(cache[currentExchange])render(currentExchange,cache[currentExchange]);};
  window.secSort  =function(col){
    if(sortCol===col)sortAsc=!sortAsc;
    else{sortCol=col;sortAsc=(col==="Name");}
    if(cache[currentExchange])render(currentExchange,cache[currentExchange]);
  };
  load("NSE");
}
EOF
