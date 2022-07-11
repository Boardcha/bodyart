
var c_seconds=parseInt(total_seconds%60);
var hours=parseInt(total_seconds/3600);
var c_minutes = Math.floor((total_seconds/60) % 60);

function CheckTime(){

	const clock = document.getElementById('clockdiv');
	const daysSpan = clock.querySelector('.days');
	const hoursSpan = clock.querySelector('.hours');
	const minutesSpan = clock.querySelector('.minutes');
	const secondsSpan = clock.querySelector('.seconds');
	
    hoursSpan.innerHTML = ('0' + hours).slice(-2);
    minutesSpan.innerHTML = ('0' + c_minutes).slice(-2);
    secondsSpan.innerHTML = ('0' + c_seconds).slice(-2);

	if(total_seconds<=0){
		document.getElementById('clockdiv').style.display = "none";
	}else{
		total_seconds=total_seconds-1;
		
		c_seconds = parseInt(total_seconds%60);
		hours = parseInt(total_seconds/3600);
		c_minutes = Math.floor((total_seconds/60) % 60);
		setTimeout("CheckTime()",1000);
	}
}
setTimeout("CheckTime()",1000);
setTimeout(function(){
	document.getElementById('countdown-timer').style.display = "block";
	document.getElementById('clockdiv').style.display = "inline-block";
}, 2000);

