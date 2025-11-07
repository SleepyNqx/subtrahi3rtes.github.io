<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>cuh</title>
<style>

body {
margin: 0;
padding: 0;
height: 100vh;
width: 100vw;
font-family: arial;
background: url(https://cdn.wallpapersafari.com/35/10/XYPfm9.jpg);
background-size: cover;
}



.WhiteButton {

transform: scale(1.9);
background-size: 2000px;
width: 100px;
height: 30px;
backdrop-filter: blur(5px)
                 brightness(111%)
                 contrast(1.5);
background: transparent;
border-radius: 200px;
border-color: transparent;
cursor: pointer;
transition: width 0.2s ease, height 0.2s ease, background 0.4s ease;
}

.WhiteButton:hover {
  width: 110px;
  height: 30px;
  background: transparent;
} 
.pos1 {
 position:relative;
 left:310px;
}

.pos2 {
 position:relative;
 right:310px;
}

.txt {
  color: #e3e3e3;
  font-size: 15px;
  filter: drop-shadow(0px 0px 0px #e3e3e3);
  transition: font-size 0.5s ease, color 0.2s ease, filter 0.2s ease;
} 

.txt:hover {
  color: white;
  font-size: 20px;
  filter: 
    drop-shadow(0px 0px 3px white);
}

.txt3 {
  color: #e3e3e3;
  font-size: 15px;
  filter: drop-shadow(0px 0px 0px #e3e3e3);
} 

.txt4 {
  color: #e3e3e3;
  font-size: 55px;
  filter: drop-shadow(2px 2px 0px #000000)
          drop-shadow(0px 0px 10px #ffffff);
} 

.txt2 {
  color: #e3e3e3;
  font-size: 35px;
  filter: drop-shadow(0px 0px 0px #e3e3e3);
  transition: font-size 0.5s ease, color 0.2s ease, filter 0.2s ease;
} 

.txt2:hover {
  color: white;
  font-size: 45px;
  filter: 
    drop-shadow(0px 0px 3px white);
}


.offset1 {
	position: relative;
	top: -30px;
	margin: 0 auto;
    display: block;
    text-align: center;
}

.offset2 {
	position: relative;
	top: 30px;
}


</style>
</head>
<body>
<div style="text-align:center;position:relative;top:200px;">

<button class="WhiteButton pos1"><span class="txt">B</span><span class="txt">I</span><span class="txt">N</span><span class="txt"> V</span><span class="txt">3</span></button>
<button class="WhiteButton"><span class="txt">n</span><span class="txt">o</span><span class="txt">w</span></button>
<button class="WhiteButton pos2">Coo?</button>

</div>
<div>
<p id="country" class="txt3 offset2">country</p>
</div>
<p id="time" class="txt4 offset1">country</p>
<script>
fetch("https://ipwho.is/")
  .then(res => res.json())
  .then(data => {
    document.getElementById("country").textContent =
      "You’re in " + data.country;
  })
  .catch(() => {
    document.getElementById("country").textContent =
      "Could not detect country.";
  });
  
function startTime() {
  const today = new Date();
  let h = today.getHours();
  let m = today.getMinutes();
  let s = today.getSeconds();
  m = checkTime(m);
  s = checkTime(s);
  document.getElementById('time').innerHTML =  h + ":" + m + ":" + s;
  setTimeout(startTime, 1000);
}

function checkTime(i) {
  if (i < 10) {i = "0" + i};  // add zero in front of numbers < 10
  return i;
}
startTime();
</script>
</body>
</html>
