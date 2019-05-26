/**
 * @author qiao / https://github.com/qiao
 * @author mrdoob / http://mrdoob.com
 * @author alteredq / http://alteredqualia.com/
 * @author WestLangley / http://github.com/WestLangley
 * @author erich666 / http://erichaines.com
 */

// This set of controls performs orbiting, dollying (zooming), and panning.
// Unlike TrackballControls, it maintains the "up" direction object.up (+Y by default).
// この一連のコントロールは、回転、回転（ズーム）、およびパンを実行します。
// TrackballControlsとは異なり、「上」方向のobject.upを維持します（デフォルトでは+ Y）。
//
//    Orbit - left mouse / touch: one finger move
//    Zoom - middle mouse, or mousewheel / touch: two finger spread or squish
//    Pan - right mouse, or arrow keys / touch: three finter swipe

// ★Orbitは「軌道」のことで、地球を周回軌道する人工衛星のようにカメラを制御する。

THREE.OrbitControls = function ( object, domElement ) {

	this.object = object;		// カメラ

	this.domElement = ( domElement !== undefined ) ? domElement : document;

	// Set to false to disable this control
	this.enabled = true;

	// "target" sets the location of focus, where the object orbits around
	// "target"は、オブジェクトが周回する焦点の位置を設定します。
	this.target = new THREE.Vector3();			// カメラの注視点(At)

	// How far you can dolly in and out ( PerspectiveCamera only )
	// ドリーインおよびアウトする距離（PerspectiveCameraのみ）
	this.minDistance = 0;
	this.maxDistance = Infinity;

	// How far you can zoom in and out ( OrthographicCamera only )
	// ズームインおよびズームアウトできる範囲（OrthographicCameraのみ）
	this.minZoom = 0;
	this.maxZoom = Infinity;

	// How far you can orbit vertically, upper and lower limits.
	// Range is 0 to Math.PI radians.
	// 上下にどのくらい遠くまで周回できるか、上限と下限。
	// 範囲は0〜Math.PIラジアンです。(180度)
	this.minPolarAngle = 0; // radians
	this.maxPolarAngle = Math.PI; // radians

	// How far you can orbit horizontally, upper and lower limits.
	// If set, must be a sub-interval of the interval [ - Math.PI, Math.PI ].
	// あなたは水平方向にどのくらい遠くまで周回できるか、上限と下限。
	// 設定した場合、区間[ -  Math.PI、Math.PI]のサブ区間でなければなりません。
	this.minAzimuthAngle = - Infinity; // radians
	this.maxAzimuthAngle = Infinity; // radians

	// Set to true to enable damping (inertia)
	// If damping is enabled, you must call controls.update() in your animation loop
	// ダンピング（慣性）を有効にするにはtrueに設定します。
	// 減衰が有効になっている場合は、アニメーションループ内でcontrols.update（）を呼び出す必要があります。
	this.enableDamping = false;
	this.dampingFactor = 0.25;

	// This option actually enables dollying in and out; left as "zoom" for backwards compatibility.
	// Set to false to disable zooming
	// このオプションは実際には出入りを許可します。 後方互換性のために "zoom"のままにしておきます。
	// ズーミングを無効にするにはfalseに設定します
	this.enableZoom = true;
	this.zoomSpeed = 1.0;

	// Set to false to disable rotating
	// 回転を無効にするにはfalseに設定します
	this.enableRotate = true;
	this.rotateSpeed = 1.0;

	// Set to false to disable panning
	// パンを無効にするには、falseに設定します。
	this.enablePan = true;
	this.keyPanSpeed = 7.0;	// pixels moved per arrow key push

	// Set to true to automatically rotate around the target
	// If auto-rotate is enabled, you must call controls.update() in your animation loop
	// ターゲットを中心に自動的に回転するには、trueに設定します。
	// 自動回転が有効になっている場合は、アニメーションループ内でcontrols.update（）を呼び出す必要があります。
	this.autoRotate = false;
	this.autoRotateSpeed = 2.0; // 30 seconds per round when fps is 60

	// Set to false to disable use of the keys
	// キーの使用を無効にするには、falseに設定します
	this.enableKeys = true;

	// The four arrow keys
	this.keys = { LEFT: 37, UP: 38, RIGHT: 39, BOTTOM: 40 };

	// Mouse buttons
	this.mouseButtons = { ORBIT: THREE.MOUSE.LEFT, ZOOM: THREE.MOUSE.MIDDLE, PAN: THREE.MOUSE.RIGHT };

	// for reset
	// リセット用
	this.target0 = this.target.clone();
	this.position0 = this.object.position.clone();
	this.zoom0 = this.object.zoom;

	//
	// public methods
	//

	// Φ 仰角 (elevation) ※上下  (Y軸からの角度)
	this.getPolarAngle = function () {

		return phi;

	};

	// θ 方位角 (azimuth) ※左右  (Z軸からの角度)
	this.getAzimuthalAngle = function () {

		return theta;

	};

	this.reset = function () {

		scope.target.copy( scope.target0 );
		scope.object.position.copy( scope.position0 );
		scope.object.zoom = scope.zoom0;

		scope.object.updateProjectionMatrix();
		scope.dispatchEvent( changeEvent );

		scope.update();

		state = STATE.NONE;

	};

	// this method is exposed, but perhaps it would be better if we can make it private...
	// このメソッドは公開されていますが、プライベートにできる方が良いかもしれません...。
	//
	// 戻り値がboolだが, クラス内では使っていないので無意味
	//
	// ★このメソッドがキモ。
	this.update = function() {

		// 視線ベクトル
		var offset = new THREE.Vector3();


		// so camera.up is the orbit axis
		// カメラです。上は軌道軸


		// クォータニオンを作成する（カメラの上方向ベクトルからY軸ベクトルへの回転）
										// このクォータニオンを、方向ベクトルvFromから方向ベクトルvToへの回転に必要な回転に設定します。
										// http://lolengine.net/blog/2013/09/18/beautiful-maths-quaternion-from-vectors
										// http://lolengine.net/blog/2014/02/24/quaternion-from-two-vectors-final
		// ⇒要は、カメラが上方向ベクトルがY軸と平行でない場合（傾いていたり）は、計算しやすくするため、計算前にY軸と平行になるように回転させておく。
		var quat = new THREE.Quaternion().setFromUnitVectors( 
											object.up, 							// vFrom  方向ベクトル   ※カメラの上方向ベクトル
											new THREE.Vector3( 0, 1, 0 ) );		// vTo    方向ベクトル	 ※基底Y軸ベクトル


		// 共役クォータニオンを作っておく
		// ⇒要は、計算終了後にY軸に回転させた分を元に戻すための逆回転用。
		var quatInverse = quat.clone().inverse();

		// 初期のカメラの上方向がデフォルト（0,1,0）であれば、quatは (w,x,y,z) = (1,0,0,0) となるので1と同じ。（共役も同じ）


		var lastPosition = new THREE.Vector3();
		var lastQuaternion = new THREE.Quaternion();


		// ここから上は1回しか呼ばれない。↑
		// --------------------------------------------------

		return function () {

			// --------------------------------------------------
			// ここから下は画面描画ごとに呼ばれる。↓


			// 視点(Eye)
			var position = scope.object.position;

			//console.log('position='+position.toLog());


			// 視点(Eye) - 注視点(At) = 視線ベクトル
			offset.copy( position ).sub( scope.target );


			// rotate offset to "y-axis-is-up" space
			// 回転オフセットを「Y軸アップ」空間で設定
			// ⇒要は「視線ベクトル」を「Y軸」合わせるために回転する（カメラの上方向が(0,1,0)以外の場合を考慮）
			offset.applyQuaternion( quat );


			// angle from z-axis around y-axis
			// Y軸を中心としたZ軸からの角度（Z軸は手前が正の右手系）を求める
			// ⇒要は「方位角(※左右)」を求める（カメラEyeがZ軸にあればthetaは0）
			theta = Math.atan2( offset.x, offset.z );	// Math.atan2(y, x)


			// angle from y-axis
			// Y軸からの角度を求める
			// ⇒要は「仰角(※上下)」を求める
			phi = Math.atan2( Math.sqrt( offset.x * offset.x + offset.z * offset.z ), offset.y );	// offset.yが何故第2引数(x)?

			//console.log('phi='+phi);



			if ( scope.autoRotate && state === STATE.NONE ) {		// autoRotateはデフォルトOFFなのでコメントアウトOK

				rotateLeft( getAutoRotationAngle() );

			}


			//console.log('thetaDelta='+thetaDelta);
			//console.log('phiDelta='+phiDelta);


			theta += thetaDelta;		// thetaDelta：0 (デフォルト)
			phi += phiDelta;			// phiDelta：0 (デフォルト)



			// restrict theta to be between desired limits
			// シータを望ましい限界の間に制限する
			//  minAzimuthAngle：制限なし(- Infinity)
			//  maxAzimuthAngle：制限なし(  Infinity)
			//  ⇒なのでコメントアウトしてもOK
			theta = Math.max( scope.minAzimuthAngle, Math.min( scope.maxAzimuthAngle, theta ) );


			// restrict phi to be between desired limits
			// ファイを望ましい限界値の範囲内に制限する
			// minPolarAngle：0
			// maxPolarAngle：Math.PI  (180度)
			// ⇒すぐあとで結局Math.PIに制限しているのでコメントアウトしてもOK (同じことやってるっぽい)
			phi = Math.max( scope.minPolarAngle, Math.min( scope.maxPolarAngle, phi ) );


			// restrict phi to be betwee EPS and PI-EPS
			// ファイをEPSとPI-EPSの間に制限する
			// EPS：0.000001
			phi = Math.max( EPS, Math.min( Math.PI - EPS, phi ) );


			// scale：1 (デフォルト)
			//  ⇒ドリー(ズーム)で変わる
			// radiusは、視線ベクトルの長さ（要はAtを球の中心としたときの半径）
			var radius = offset.length() * scale;

			//console.log('radius='+radius);


			// restrict radius to be between desired limits
			// 半径を制限範囲内に制限する
			// minDistance：0
			// maxDistance：制限なし(  Infinity)
			// ⇒コメントアウトしてもOK
			radius = Math.max( scope.minDistance, Math.min( scope.maxDistance, radius ) );


			// move target to panned location
			// ターゲットを画面移動した位置に移動する
			// ⇒要はカメラの注視点(At)をパンする
			scope.target.add( panOffset );

			// 球面座標系から直交座標系(右手)へ変換した座標を視線ベクトルとする
			//  ⇒https://qiita.com/edo_m18/items/d80171fcada047c454b9
			//   ⇒ここの解説は納得したがThreeJSと式が違う。⇒ThreeJSが正しい。
			offset.x = radius * Math.sin( phi ) * Math.sin( theta );	// thetaはZ軸から角度なので、X座標はsin(theta)。これはY座標にも比例するのでsin(phi)
			offset.y = radius * Math.cos( phi );						// phiはY軸から角度なので、Y座標はcos(phi)
			offset.z = radius * Math.sin( phi ) * Math.cos( theta );	// thetaはZ軸から角度なので、Z座標はcos(theta)。これはY座標にも比例するのでsin(phi)


			// rotate offset back to "camera-up-vector-is-up" space
			// 視線ベクトルを「カメラの上方向ベクトル」空間に戻すために共役クォータニオンで逆回転させる。
			// カメラの上方向がデフォルト(0,1,0)のままならコメントアウトOK
			offset.applyQuaternion( quatInverse );


			// 注視点ベクトル＋視線ベクトル＝カメラ位置ベクトル
			position.copy( scope.target ).add( offset );


			// ビュー行列を更新
			scope.object.lookAt( scope.target );


			// 終わり


			if ( scope.enableDamping === true ) {		// 慣性モード?   enableDamping：false (デフォルト)なのでコメントアウトOK

				thetaDelta *= ( 1 - scope.dampingFactor );
				phiDelta *= ( 1 - scope.dampingFactor );

			} else {

				// 0に戻さないと回転し続ける
				thetaDelta = 0;
				phiDelta = 0;

			}

			scale = 1;
			panOffset.set( 0, 0, 0 );

			// update condition is:
			// min(camera displacement, camera rotation in radians)^2 > EPS
			// using small-angle approximation cos(x/2) = 1 - x^2 / 8
			// 更新条件:
			// min(カメラ変位、ラジアン単位のカメラ回転)^2 > EPS
			// 小角近似 cos(x/2) = 1 - x^2 / 8 を使用


			// コメントアウトしても問題なし(kawa)
			/*

			if ( zoomChanged ||
				lastPosition.distanceToSquared( scope.object.position ) > EPS ||
				8 * ( 1 - lastQuaternion.dot( scope.object.quaternion ) ) > EPS ) {

				scope.dispatchEvent( changeEvent );

				lastPosition.copy( scope.object.position );
				lastQuaternion.copy( scope.object.quaternion );
				zoomChanged = false;

				return true;

			}

			return false;

			*/

		};

	}();

	this.dispose = function() {

		scope.domElement.removeEventListener( 'contextmenu', onContextMenu, false );
		scope.domElement.removeEventListener( 'mousedown', onMouseDown, false );
		scope.domElement.removeEventListener( 'mousewheel', onMouseWheel, false );
		scope.domElement.removeEventListener( 'MozMousePixelScroll', onMouseWheel, false ); // firefox

		scope.domElement.removeEventListener( 'touchstart', onTouchStart, false );
		scope.domElement.removeEventListener( 'touchend', onTouchEnd, false );
		scope.domElement.removeEventListener( 'touchmove', onTouchMove, false );

		document.removeEventListener( 'mousemove', onMouseMove, false );
		document.removeEventListener( 'mouseup', onMouseUp, false );
		document.removeEventListener( 'mouseout', onMouseUp, false );

		window.removeEventListener( 'keydown', onKeyDown, false );

		//scope.dispatchEvent( { type: 'dispose' } ); // should this be added here?

	};

	//
	// internals
	//

	var scope = this;		// thisはカメラオブジェクト

	var changeEvent = { type: 'change' };
	var startEvent = { type: 'start' };
	var endEvent = { type: 'end' };

	var STATE = { NONE : - 1, ROTATE : 0, DOLLY : 1, PAN : 2, TOUCH_ROTATE : 3, TOUCH_DOLLY : 4, TOUCH_PAN : 5 };

	var state = STATE.NONE;

	var EPS = 0.000001;		// EPS = ε イプシロン(小さい数を表すために数学ではよく使用される)

	// current position in spherical coordinates
	// 「球面座標」での現在位置
	//  ⇒これがキモ。球面座標ではある1点の座標を(r, θ, Φ)で表す。
	//    rは球の中心からの半径。つまり「カメラの視点と注視点のベクトル」
	var theta;		// θ 方位角 (azimuth) ※左右  (Z軸からの角度)
	var phi;		// Φ 仰角 (elevation) ※上下  (Y軸からの角度)

	var phiDelta = 0;
	var thetaDelta = 0;
	var scale = 1;
	var panOffset = new THREE.Vector3();
	var zoomChanged = false;

	var rotateStart = new THREE.Vector2();
	var rotateEnd = new THREE.Vector2();
	var rotateDelta = new THREE.Vector2();

	var panStart = new THREE.Vector2();
	var panEnd = new THREE.Vector2();
	var panDelta = new THREE.Vector2();

	var dollyStart = new THREE.Vector2();
	var dollyEnd = new THREE.Vector2();
	var dollyDelta = new THREE.Vector2();

	function getAutoRotationAngle() {

		return 2 * Math.PI / 60 / 60 * scope.autoRotateSpeed;

	}

	// ★
	function getZoomScale() {

		return Math.pow( 0.95, scope.zoomSpeed );

	}

	// ★
	// 左右方向の回転
	function rotateLeft( angle ) {

		thetaDelta -= angle;	// 符号を変えるの回転方向が反転（直感的では無くなる）

	}

	// ★
	// 上限方向の回転
	function rotateUp( angle ) {

		phiDelta -= angle;	// 符号を変えるの回転方向が反転（直感的では無くなる）

	}

	// ★
	var panLeft = function() {

		var v = new THREE.Vector3();

		return function panLeft( distance, objectMatrix ) {

			var te = objectMatrix.elements;

			// get X column of objectMatrix
			v.set( te[ 0 ], te[ 1 ], te[ 2 ] );

			v.multiplyScalar( - distance );

			panOffset.add( v );

		};

	}();

	// ★
	var panUp = function() {

		var v = new THREE.Vector3();

		return function panUp( distance, objectMatrix ) {

			var te = objectMatrix.elements;

			// get Y column of objectMatrix
			v.set( te[ 4 ], te[ 5 ], te[ 6 ] );

			v.multiplyScalar( distance );

			panOffset.add( v );

		};

	}();

	// ★
	// deltaX and deltaY are in pixels; right and down are positive
	// deltaXおよびdeltaYはピクセル単位;右と下が正である
	var pan = function() {

		var offset = new THREE.Vector3();

		return function( deltaX, deltaY ) {

			var element = scope.domElement === document ? scope.domElement.body : scope.domElement;

			if ( scope.object instanceof THREE.PerspectiveCamera ) {

				// perspective
				var position = scope.object.position;
				offset.copy( position ).sub( scope.target );
				var targetDistance = offset.length();

				// half of the fov is center to top of screen
				// FOVの半分は画面の中央から上
				targetDistance *= Math.tan( ( scope.object.fov / 2 ) * Math.PI / 180.0 );

				// we actually don't use screenWidth, since perspective camera is fixed to screen height
				// 遠近法カメラは画面の高さに固定されているので、実際にはscreenWidthを使用しません。
				panLeft( 2 * deltaX * targetDistance / element.clientHeight, scope.object.matrix );
				panUp( 2 * deltaY * targetDistance / element.clientHeight, scope.object.matrix );

			} else if ( scope.object instanceof THREE.OrthographicCamera ) {

				// orthographic
				panLeft( deltaX * ( scope.object.right - scope.object.left ) / element.clientWidth, scope.object.matrix );
				panUp( deltaY * ( scope.object.top - scope.object.bottom ) / element.clientHeight, scope.object.matrix );

			} else {

				// camera neither orthographic nor perspective
				console.warn( 'WARNING: OrbitControls.js encountered an unknown camera type - pan disabled.' );
				scope.enablePan = false;

			}

		};

	}();

	// ★
	function dollyIn( dollyScale ) {

		if ( scope.object instanceof THREE.PerspectiveCamera ) {

			scale /= dollyScale;

		} else if ( scope.object instanceof THREE.OrthographicCamera ) {

			scope.object.zoom = Math.max( scope.minZoom, Math.min( scope.maxZoom, scope.object.zoom * dollyScale ) );
			scope.object.updateProjectionMatrix();
			zoomChanged = true;

		} else {

			console.warn( 'WARNING: OrbitControls.js encountered an unknown camera type - dolly/zoom disabled.' );
			scope.enableZoom = false;

		}

	}

	// ★
	function dollyOut( dollyScale ) {

		if ( scope.object instanceof THREE.PerspectiveCamera ) {

			scale *= dollyScale;

		} else if ( scope.object instanceof THREE.OrthographicCamera ) {

			scope.object.zoom = Math.max( scope.minZoom, Math.min( scope.maxZoom, scope.object.zoom / dollyScale ) );
			scope.object.updateProjectionMatrix();
			zoomChanged = true;

		} else {

			console.warn( 'WARNING: OrbitControls.js encountered an unknown camera type - dolly/zoom disabled.' );
			scope.enableZoom = false;

		}

	}

	//
	// event callbacks - update the object state
	//

	// ★
	// 回転中のマウスダウンからのみ呼ばれる
	function handleMouseDownRotate( event ) {

		//console.log( 'handleMouseDownRotate' );

		rotateStart.set( event.clientX, event.clientY );

	}

	// ★
	function handleMouseDownDolly( event ) {

		//console.log( 'handleMouseDownDolly' );

		dollyStart.set( event.clientX, event.clientY );

	}

	// ★
	function handleMouseDownPan( event ) {

		//console.log( 'handleMouseDownPan' );

		panStart.set( event.clientX, event.clientY );

	}

	// ★
	// 回転中のマウスムーブからのみ呼ばれる
	function handleMouseMoveRotate( event ) {

		//console.log( 'handleMouseMoveRotate' );

		rotateEnd.set( event.clientX, event.clientY );


		// rotateStartはマウスダウンで設定される
		rotateDelta.subVectors( rotateEnd, rotateStart );


		var element = scope.domElement === document ? scope.domElement.body : scope.domElement;


		// rotating across whole screen goes 360 degrees around
		// 画面全体を360度回転
		// 2 * Math.PI = 360度
		//
		// 要は左右方向の回転角度を求める
		rotateLeft( 2 * Math.PI * rotateDelta.x / element.clientWidth * scope.rotateSpeed );


		// rotating up and down along whole screen attempts to go 360, but limited to 180
		// 画面全体に沿って上下に回転させると360度に移動しようとしますが、180度に制限されます。
		// 
		// 要は上下方向の回転角度を求める
		rotateUp( 2 * Math.PI * rotateDelta.y / element.clientHeight * scope.rotateSpeed );


		rotateStart.copy( rotateEnd );


		scope.update();

	}

	// ★
	function handleMouseMoveDolly( event ) {

		//console.log( 'handleMouseMoveDolly' );

		dollyEnd.set( event.clientX, event.clientY );

		dollyDelta.subVectors( dollyEnd, dollyStart );

		if ( dollyDelta.y > 0 ) {

			dollyIn( getZoomScale() );

		} else if ( dollyDelta.y < 0 ) {

			dollyOut( getZoomScale() );

		}

		dollyStart.copy( dollyEnd );

		scope.update();

	}

	// ★
	function handleMouseMovePan( event ) {

		//console.log( 'handleMouseMovePan' );

		panEnd.set( event.clientX, event.clientY );

		panDelta.subVectors( panEnd, panStart );

		pan( panDelta.x, panDelta.y );

		panStart.copy( panEnd );

		scope.update();

	}

	function handleMouseUp( event ) {

		//console.log( 'handleMouseUp' );

	}

	// ★
	function handleMouseWheel( event ) {

		//console.log( 'handleMouseWheel' );

		var delta = 0;

		if ( event.wheelDelta !== undefined ) {

			// WebKit / Opera / Explorer 9

			delta = event.wheelDelta;

		} else if ( event.detail !== undefined ) {

			// Firefox

			delta = - event.detail;

		}

		if ( delta > 0 ) {

			dollyOut( getZoomScale() );

		} else if ( delta < 0 ) {

			dollyIn( getZoomScale() );

		}

		scope.update();

	}

	function handleKeyDown( event ) {

		//console.log( 'handleKeyDown' );

		switch ( event.keyCode ) {

			case scope.keys.UP:
				pan( 0, scope.keyPanSpeed );
				scope.update();
				break;

			case scope.keys.BOTTOM:
				pan( 0, - scope.keyPanSpeed );
				scope.update();
				break;

			case scope.keys.LEFT:
				pan( scope.keyPanSpeed, 0 );
				scope.update();
				break;

			case scope.keys.RIGHT:
				pan( - scope.keyPanSpeed, 0 );
				scope.update();
				break;

		}

	}

	function handleTouchStartRotate( event ) {

		//console.log( 'handleTouchStartRotate' );

		rotateStart.set( event.touches[ 0 ].pageX, event.touches[ 0 ].pageY );

	}

	function handleTouchStartDolly( event ) {

		//console.log( 'handleTouchStartDolly' );

		var dx = event.touches[ 0 ].pageX - event.touches[ 1 ].pageX;
		var dy = event.touches[ 0 ].pageY - event.touches[ 1 ].pageY;

		var distance = Math.sqrt( dx * dx + dy * dy );

		dollyStart.set( 0, distance );

	}

	function handleTouchStartPan( event ) {

		//console.log( 'handleTouchStartPan' );

		panStart.set( event.touches[ 0 ].pageX, event.touches[ 0 ].pageY );

	}

	function handleTouchMoveRotate( event ) {

		//console.log( 'handleTouchMoveRotate' );

		rotateEnd.set( event.touches[ 0 ].pageX, event.touches[ 0 ].pageY );
		rotateDelta.subVectors( rotateEnd, rotateStart );

		var element = scope.domElement === document ? scope.domElement.body : scope.domElement;

		// rotating across whole screen goes 360 degrees around
		rotateLeft( 2 * Math.PI * rotateDelta.x / element.clientWidth * scope.rotateSpeed );

		// rotating up and down along whole screen attempts to go 360, but limited to 180
		rotateUp( 2 * Math.PI * rotateDelta.y / element.clientHeight * scope.rotateSpeed );

		rotateStart.copy( rotateEnd );

		scope.update();

	}

	function handleTouchMoveDolly( event ) {

		//console.log( 'handleTouchMoveDolly' );

		var dx = event.touches[ 0 ].pageX - event.touches[ 1 ].pageX;
		var dy = event.touches[ 0 ].pageY - event.touches[ 1 ].pageY;

		var distance = Math.sqrt( dx * dx + dy * dy );

		dollyEnd.set( 0, distance );

		dollyDelta.subVectors( dollyEnd, dollyStart );

		if ( dollyDelta.y > 0 ) {

			dollyOut( getZoomScale() );

		} else if ( dollyDelta.y < 0 ) {

			dollyIn( getZoomScale() );

		}

		dollyStart.copy( dollyEnd );

		scope.update();

	}

	function handleTouchMovePan( event ) {

		//console.log( 'handleTouchMovePan' );

		panEnd.set( event.touches[ 0 ].pageX, event.touches[ 0 ].pageY );

		panDelta.subVectors( panEnd, panStart );

		pan( panDelta.x, panDelta.y );

		panStart.copy( panEnd );

		scope.update();

	}

	function handleTouchEnd( event ) {

		//console.log( 'handleTouchEnd' );

	}

	//
	// event handlers - FSM: listen for events and reset state
	//

	// ★
	function onMouseDown( event ) {

		if ( scope.enabled === false ) return;

		event.preventDefault();

		if ( event.button === scope.mouseButtons.ORBIT ) {

			if ( scope.enableRotate === false ) return;

			handleMouseDownRotate( event );

			state = STATE.ROTATE;

		} else if ( event.button === scope.mouseButtons.ZOOM ) {

			if ( scope.enableZoom === false ) return;

			handleMouseDownDolly( event );

			state = STATE.DOLLY;

		} else if ( event.button === scope.mouseButtons.PAN ) {

			if ( scope.enablePan === false ) return;

			handleMouseDownPan( event );

			state = STATE.PAN;

		}

		if ( state !== STATE.NONE ) {

			document.addEventListener( 'mousemove', onMouseMove, false );
			document.addEventListener( 'mouseup', onMouseUp, false );
			document.addEventListener( 'mouseout', onMouseUp, false );

			scope.dispatchEvent( startEvent );

		}

	}

	// ★
	function onMouseMove( event ) {

		if ( scope.enabled === false ) return;

		event.preventDefault();

		if ( state === STATE.ROTATE ) {

			if ( scope.enableRotate === false ) return;

			handleMouseMoveRotate( event );

		} else if ( state === STATE.DOLLY ) {

			if ( scope.enableZoom === false ) return;

			handleMouseMoveDolly( event );

		} else if ( state === STATE.PAN ) {

			if ( scope.enablePan === false ) return;

			handleMouseMovePan( event );

		}

	}

	function onMouseUp( event ) {

		if ( scope.enabled === false ) return;

		// handleMouseUpでは何もしていない
		handleMouseUp( event );

		document.removeEventListener( 'mousemove', onMouseMove, false );
		document.removeEventListener( 'mouseup', onMouseUp, false );
		document.removeEventListener( 'mouseout', onMouseUp, false );

		scope.dispatchEvent( endEvent );

		state = STATE.NONE;

	}

	// ★
	function onMouseWheel( event ) {

		if ( scope.enabled === false || scope.enableZoom === false || state !== STATE.NONE ) return;

		event.preventDefault();
		event.stopPropagation();

		handleMouseWheel( event );

		scope.dispatchEvent( startEvent ); // not sure why these are here...
		scope.dispatchEvent( endEvent );

	}

	function onKeyDown( event ) {

		if ( scope.enabled === false || scope.enableKeys === false || scope.enablePan === false ) return;

		handleKeyDown( event );

	}

	function onTouchStart( event ) {

		if ( scope.enabled === false ) return;

		switch ( event.touches.length ) {

			case 1:	// one-fingered touch: rotate

				if ( scope.enableRotate === false ) return;

				handleTouchStartRotate( event );

				state = STATE.TOUCH_ROTATE;

				break;

			case 2:	// two-fingered touch: dolly

				if ( scope.enableZoom === false ) return;

				handleTouchStartDolly( event );

				state = STATE.TOUCH_DOLLY;

				break;

			case 3: // three-fingered touch: pan

				if ( scope.enablePan === false ) return;

				handleTouchStartPan( event );

				state = STATE.TOUCH_PAN;

				break;

			default:

				state = STATE.NONE;

		}

		if ( state !== STATE.NONE ) {

			scope.dispatchEvent( startEvent );

		}

	}

	function onTouchMove( event ) {

		if ( scope.enabled === false ) return;

		event.preventDefault();
		event.stopPropagation();

		switch ( event.touches.length ) {

			case 1: // one-fingered touch: rotate

				if ( scope.enableRotate === false ) return;
				if ( state !== STATE.TOUCH_ROTATE ) return; // is this needed?...

				handleTouchMoveRotate( event );

				break;

			case 2: // two-fingered touch: dolly

				if ( scope.enableZoom === false ) return;
				if ( state !== STATE.TOUCH_DOLLY ) return; // is this needed?...

				handleTouchMoveDolly( event );

				break;

			case 3: // three-fingered touch: pan

				if ( scope.enablePan === false ) return;
				if ( state !== STATE.TOUCH_PAN ) return; // is this needed?...

				handleTouchMovePan( event );

				break;

			default:

				state = STATE.NONE;

		}

	}

	function onTouchEnd( event ) {

		if ( scope.enabled === false ) return;

		handleTouchEnd( event );

		scope.dispatchEvent( endEvent );

		state = STATE.NONE;

	}

	function onContextMenu( event ) {

		event.preventDefault();

	}

	//

	scope.domElement.addEventListener( 'contextmenu', onContextMenu, false );

	scope.domElement.addEventListener( 'mousedown', onMouseDown, false );
	scope.domElement.addEventListener( 'mousewheel', onMouseWheel, false );
	scope.domElement.addEventListener( 'MozMousePixelScroll', onMouseWheel, false ); // firefox

	scope.domElement.addEventListener( 'touchstart', onTouchStart, false );
	scope.domElement.addEventListener( 'touchend', onTouchEnd, false );
	scope.domElement.addEventListener( 'touchmove', onTouchMove, false );

	window.addEventListener( 'keydown', onKeyDown, false );

	// force an update at start

	this.update();

};

THREE.OrbitControls.prototype = Object.create( THREE.EventDispatcher.prototype );
THREE.OrbitControls.prototype.constructor = THREE.OrbitControls;

Object.defineProperties( THREE.OrbitControls.prototype, {

	center: {

		get: function () {

			console.warn( 'THREE.OrbitControls: .center has been renamed to .target' );
			return this.target;

		}

	},

	// backward compatibility

	noZoom: {

		get: function () {

			console.warn( 'THREE.OrbitControls: .noZoom has been deprecated. Use .enableZoom instead.' );
			return ! this.enableZoom;

		},

		set: function ( value ) {

			console.warn( 'THREE.OrbitControls: .noZoom has been deprecated. Use .enableZoom instead.' );
			this.enableZoom = ! value;

		}

	},

	noRotate: {

		get: function () {

			console.warn( 'THREE.OrbitControls: .noRotate has been deprecated. Use .enableRotate instead.' );
			return ! this.enableRotate;

		},

		set: function ( value ) {

			console.warn( 'THREE.OrbitControls: .noRotate has been deprecated. Use .enableRotate instead.' );
			this.enableRotate = ! value;

		}

	},

	noPan: {

		get: function () {

			console.warn( 'THREE.OrbitControls: .noPan has been deprecated. Use .enablePan instead.' );
			return ! this.enablePan;

		},

		set: function ( value ) {

			console.warn( 'THREE.OrbitControls: .noPan has been deprecated. Use .enablePan instead.' );
			this.enablePan = ! value;

		}

	},

	noKeys: {

		get: function () {

			console.warn( 'THREE.OrbitControls: .noKeys has been deprecated. Use .enableKeys instead.' );
			return ! this.enableKeys;

		},

		set: function ( value ) {

			console.warn( 'THREE.OrbitControls: .noKeys has been deprecated. Use .enableKeys instead.' );
			this.enableKeys = ! value;

		}

	},

	staticMoving : {

		get: function () {

			console.warn( 'THREE.OrbitControls: .staticMoving has been deprecated. Use .enableDamping instead.' );
			return ! this.constraint.enableDamping;

		},

		set: function ( value ) {

			console.warn( 'THREE.OrbitControls: .staticMoving has been deprecated. Use .enableDamping instead.' );
			this.constraint.enableDamping = ! value;

		}

	},

	dynamicDampingFactor : {

		get: function () {

			console.warn( 'THREE.OrbitControls: .dynamicDampingFactor has been renamed. Use .dampingFactor instead.' );
			return this.constraint.dampingFactor;

		},

		set: function ( value ) {

			console.warn( 'THREE.OrbitControls: .dynamicDampingFactor has been renamed. Use .dampingFactor instead.' );
			this.constraint.dampingFactor = value;

		}

	}

} );
