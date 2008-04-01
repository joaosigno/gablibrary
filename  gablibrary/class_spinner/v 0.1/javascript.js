var m_step;
var m_val_min;
var m_val_max;
var m_looping;

function init(step, val_min, val_max, looping)
{
	m_step 		= parseFloat(step);
	m_val_min 	= val_min;
	m_val_max 	= val_max;
	if(looping.toLowerCase() == 'true')
		m_looping	= true;
	else
		m_looping	= false;
}

function shiftUpDown(obj, step, val_min, val_max, looping)
{	
	checkCorrectInput(obj, val_min, val_max);
	if(event.wheelDelta < 0)
	{
		init(event.wheelDelta/120*(-1)*step, val_min, val_max, looping);
		count('down', obj);
	}
	else
	{
		init(event.wheelDelta/120*step, val_min, val_max, looping);
		count('up', obj);
	}
	
	return false;
}


function callCount(val, obj, step, val_min, val_max, looping)
{
	init(step, val_min, val_max, looping);
	currentKeyEvent = event.keyCode;

	switch(currentKeyEvent)
	{
		// --- buttons ---
		case 0:
			count(val, obj);
			break;
		// --- up ---
		case 38:
			count('up', obj);
			break;
		// --- down ---
		case 40:
			count('down', obj);
			break;
		// --- pageUp ---
		case 33:
			obj.value = val_max;
			changeState(obj, "disable_increase");
			changeState(obj, "enable_decrease");
			break;
		// --- pageDown ---
		case 34:
			obj.value = val_min;
			changeState(obj, "disable_decrease");
			changeState(obj, "enable_increase");
			break;
	}
}

function changeState(obj, mode)
{
	eval("var orient = vdo" + obj.id);

	if(eval("vdCustomMode_" + obj.id + " == 1"))
	{
		var customValue;
		if(event.keyCode == 33)
			customValue = m_val_max;
		else if(event.keyCode == 34)
			customValue = m_val_min;
				
		eval("vdArrayCount_" + obj.id + " = " + customValue);
		eval("obj.value = vdCustomValues_" + obj.id + "[" + customValue + "]");
	}
		
	if(m_looping)
		return;
		
	decreaseButton = document.getElementById("down_" + obj.id);
	increaseButton = document.getElementById("up_" + obj.id);
	
	switch(mode)
	{
		case "enable_decrease":
			imgButton = document.getElementById("img" + obj.id + "_down");
			imgButton.src = (orient == 1) ? m_Images[2] : m_Images[0];
			decreaseButton.disabled = false;
			break;
			
		case "disable_increase":
			imgButton = document.getElementById("img" + obj.id + "_up");
			imgButton.src = (orient == 1) ? m_Images[5] : m_Images[7];
			increaseButton.disabled = true;
			break;
			
		case "enable_increase":
			imgButton = document.getElementById("img" + obj.id + "_up");
			imgButton.src = (orient == 1) ? m_Images[4] : m_Images[6];
			increaseButton.disabled = false;
			break;
			
		case "disable_decrease":
			imgButton = document.getElementById("img" + obj.id + "_down");
			imgButton.src = (orient == 1) ? m_Images[3] : m_Images[1];
			decreaseButton.disabled = true;
			break;
	}
}

function count(val, obj)
{
	var tmp = parseFloat(obj.value);

	if(eval("vdCustomMode_" + obj.id + " == 1")) // customControl -> use variable to count up and down
		eval("tmp = vdArrayCount_" + obj.id);

	eval("var orient = vdo" + obj.id);
	
	if(val == 'up')
	{
		if((m_looping) && (tmp*1 + m_step > m_val_max))	// looping allowed and maximum reached -> set back to minimum
			tmp = m_val_min;
		else
		{
			if((m_val_max == 0) && (m_val_min == 0)) // no min and max values (both are 0)
				tmp = tmp * 1 + m_step;
			else
				if(tmp*1 + m_step < m_val_max) 	// maximum not reached until now
				{
					tmp = tmp * 1 + m_step;
					imgButton = document.getElementById("down_" + obj.id);
					if(imgButton.disabled)
						changeState(obj, "enable_decrease");
				}
				else
				{ 
					tmp = m_val_max; // maximum will be reached -> set value to max
					changeState(obj, "disable_increase");
				}
		}
	}
	else
	{
		if((m_looping) && (tmp - m_step < m_val_min)) // looping allowed and minimum reached -> set to maximum
			tmp = m_val_max;
		else
		{
			if((m_val_max == 0) && (m_val_min == 0)) // no min and max values (both are 0)
				tmp -= m_step;
			else
				if(tmp - m_step > m_val_min) // minimum not reached until now
				{
					tmp -= m_step;
					imgButton = document.getElementById("up_" + obj.id);
					if(imgButton.disabled)
						changeState(obj, "enable_increase");
				}
				else
				{
					tmp = m_val_min; // minimum will be reached -> set value to min
					changeState(obj, "disable_decrease");
				}
		}		
	}

	if(eval("vdi" + obj.id + " != 0")) // fill up integer places
		tmp = fillUp(tmp, eval("vdi" + obj.id));
	
	if(eval("vdd" + obj.id + " != 0")) // fill up decimal places and round it
		tmp = roundIt(tmp, eval("vdd" + obj.id)*(-1));
	
	if(eval("vdCustomMode_" + obj.id + " == 1")) // customControl -> use array values
	{
		eval("vdArrayCount_" + obj.id + " = tmp");
		eval("tmp = vdCustomValues_" + obj.id + "[tmp]");
	}
	
	obj.value = tmp;
}

function checkCorrectInput(obj, val_min, val_max)
{
	var tmp = parseFloat(obj.value);
	if(isNaN(tmp))
		if(eval("vdCustomMode_" + obj.id + " != 1"))
			obj.value = val_min;
	if(!((val_min == 0) && (val_max == 0)))
	{
		if(tmp > val_max)
			obj.value = val_max;
		if(tmp < val_min)
			obj.value = val_min;
	}
}

function fillUp(digit, mval)
{
	var tmp = digit.toString();
	var places = "";
	if(tmp.length >= mval)
		return digit;
		
	for(i = 0; i < mval; i++)
		places += "0";
	
	tmp = places + tmp;
	tmp = tmp.substring(tmp.length-mval);
	return tmp;
}

function roundIt(digit,mval)
{
	/*Maerz 2000, copyright Antje Hofmann mail:ah@pc-anfaenger.de*/
	if (typeof(digit) == "string")
		if (digit.indexOf(",") != -1)
			digit = digit.substring(0,digit.indexOf(","))+"." + digit.substring(digit.indexOf(",")+1,digit.length)
	
	digit = Math.round(digit/Math.pow(10,mval))*Math.pow(10,mval);
	
	digit = digit + "";
	if (digit.indexOf(".")!=-1)
		if (digit.length-digit.indexOf(".")>Math.abs(mval)+1)
			digit = digit.substring(0,digit.indexOf(".")+Math.abs(mval)+1);
	
	if (mval<0)
	{
		if (digit.indexOf(".") == -1) 
			digit = digit + ".";
		if (digit.indexOf(".") == 0) 
			digit = "0" + digit;
		while (Math.abs(mval)-(digit.length-digit.indexOf("."))>-1)
			digit = digit + "0";
	}
	
	return digit;
}
