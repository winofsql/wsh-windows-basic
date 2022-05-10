// *************************************
// ウインドウの位置とサイズ
// *************************************
function baseWindow( x, y, w, h ) {

	if ( (location.href).substr(0,4) == "file" ) {

		top.moveTo( x, y );
		top.resizeTo( w, h );

	}

}

// *************************************
// デスクトップ中央
// *************************************
function centerWindow( w, h ) {

	if ( (location.href).substr(0,4) == "file" ) {

		// ウインドウの位置とサイズ
		top.resizeTo( w, h );
		top.moveTo((screen.width-w)/2, (screen.height-h)/2 )

	}

}

// *************************************
// CreateObject
// *************************************
function newObject( className ) {

	var obj;

	try {
		obj = new ActiveXObject( className );
	}
	catch (e) {
		obj = null;
	}

	return obj;

}
