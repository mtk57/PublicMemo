/**
 * @author Eberhard Graether / http://egraether.com/
 * @author Mark Lundin 	/ http://mark-lundin.com
 * @author Simone Manini / http://daron1337.github.io
 * @author Luca Antiga 	/ http://lantiga.github.io
 */

// object = Camera
THREE.TrackballControls = function ( object, domElement ) {

	var _this = this;
	var STATE = { NONE: - 1, ROTATE: 0, ZOOM: 1, PAN: 2, TOUCH_ROTATE: 3, TOUCH_ZOOM_PAN: 4 };


	this.object = object;		// object = Camera
	this.domElement = ( domElement !== undefined ) ? domElement : document;

	// API

	this.enabled = true;

	this.screen = { left: 0, top: 0, width: 0, height: 0 };

	this.rotateSpeed = 1.0;				// 回転速度(Default=1.0)
	this.zoomSpeed = 2.0;				// ズーム速度(Default=1.2)
	this.panSpeed = 0.3;				// パン速度(Default=0.3)

	this.noRotate = false;
	this.noZoom = false;
	this.noPan = false;

	this.staticMoving = true;						// 余韻(?)の有無(false=余韻あり(デフォルト))
	this.dynamicDampingFactor = 0.2;  				// 小さくすると、余韻(?)が長く続く(デフォルト=0.2。staticMovingがfalseのときに有効)

	this.minDistance = 0;
	this.maxDistance = Infinity;

	this.keys = [ 65 /*A*/, 83 /*S*/, 68 /*D*/ ];

	// internals

	this.target = new THREE.Vector3();

	var EPS = 0.000001;

	var lastPosition = new THREE.Vector3();

	var _state = STATE.NONE,
	_prevState = STATE.NONE,

	_eye = new THREE.Vector3(),				// カメラ位置

	_movePrev = new THREE.Vector2(),		// 前のマウス座標(正規化)
	_moveCurr = new THREE.Vector2(),		// 現在のマウス座標(正規化)

	_lastAxis = new THREE.Vector3(),		// 回転後の余韻用の回転軸
	_lastAngle = 0,							// 回転後の余韻用の回転角度(Radian)

	_zoomStart = new THREE.Vector2(),
	_zoomEnd = new THREE.Vector2(),

	_touchZoomDistanceStart = 0,
	_touchZoomDistanceEnd = 0,

	_panStart = new THREE.Vector2(),
	_panEnd = new THREE.Vector2();

	// for reset

	this.target0 = this.target.clone();
	this.position0 = this.object.position.clone();
	this.up0 = this.object.up.clone();

	// events

	var changeEvent = { type: 'change' };
	var startEvent = { type: 'start' };
	var endEvent = { type: 'end' };


	// methods

	this.handleResize = function () {

		if ( this.domElement === document ) {

			this.screen.left = 0;
			this.screen.top = 0;
			this.screen.width = window.innerWidth;
			this.screen.height = window.innerHeight;

		} else {

			var box = this.domElement.getBoundingClientRect();
			// adjustments come from similar code in the jquery offset() function
			var d = this.domElement.ownerDocument.documentElement;
			this.screen.left = box.left + window.pageXOffset - d.clientLeft;
			this.screen.top = box.top + window.pageYOffset - d.clientTop;
			this.screen.width = box.width;
			this.screen.height = box.height;

		}

	};

	this.handleEvent = function ( event ) {

		if ( typeof this[ event.type ] == 'function' ) {

			this[ event.type ]( event );

		}

	};

	// スクリーン座標を左上原点で正規化した座標。0～1
	//
	// pageX, pageYがスクリーン座標
	// _this.screen.left, _this.screen.topは 0
	//
	// 戻り値のvectorには、左上を原点とした「正規化されたスクリーン座標」が返る。(つまり画面サイズに関係なく、0～1になる)
	var getMouseOnScreen = ( function () {

		var vector = new THREE.Vector2();

		return function getMouseOnScreen( pageX, pageY ) {

			vector.set(
				( pageX - _this.screen.left ) / _this.screen.width,
				( pageY - _this.screen.top ) / _this.screen.height
			);

			return vector;

		};

	}() );

	var getMouseOnCircle = ( function () {

		var vector = new THREE.Vector2();

		return function getMouseOnCircle( pageX, pageY ) {

			// kawa add >>
			//console.log("getMouseOnCircle  pageX,pageY="+pageX+", "+pageY);
			//console.log("screenW="+_this.screen.width+", screenL="+_this.screen.left);
			//console.log("screenH="+_this.screen.height+", screenT="+_this.screen.top);

			
			// マウス位置の座標を正規化された2次元ベクトルに変換します
			// 中心 0, 0  右上 1, 1  左下 -1, -1
			// kawa add <<

			vector.set(
				( ( pageX - _this.screen.width * 0.5 - _this.screen.left ) / ( _this.screen.width * 0.5 ) ),
				( ( _this.screen.height + 2 * ( _this.screen.top - pageY ) ) / _this.screen.width ) // screen.width intentional
			);

			//console.log("vector="+vector.x+", "+vector.y);	// kawa add

			return vector;

		};

	}() );

	// カメラ回転
	this.rotateCamera = ( function() {

		var axis = new THREE.Vector3(),
			quaternion = new THREE.Quaternion(),
			eyeDirection = new THREE.Vector3(),
			objectUpDirection = new THREE.Vector3(),
			objectSidewaysDirection = new THREE.Vector3(),
			moveDirection = new THREE.Vector3(),
			angle;

		// kawa chg >>
		//return function rotateCamera() {
		return function rotateCamera(arrowHelper) {
		// kawa chg <<

			// マウス移動方向ベクトル
			//  ⇒マウス移動量の差分
			//  ⇒マウス座標は正規化済
			moveDirection.set( _moveCurr.x - _movePrev.x, _moveCurr.y - _movePrev.y, 0 );


			// 回転角度
			//  ⇒マウス移動方向ベクトルの長さ
			//  ⇒単位円の直角三角形の高さのイメージ
			//  ⇒ラジアンの近似と思われる
			//  ⇒必ず1未満になる
			angle = moveDirection.length();





			if ( angle ) {

				//console.log('_this.object.position(1)='+_this.object.position.x + ', ' + this.object.position.y + ', ' + this.object.position.z);

				// 注視点⇒視点ベクトル
				//  ⇒(現在の視点 - 注視点)     ※注視点の初期値は原点。PANで更新。
				//  ⇒object == Camera
				//  ⇒_eye == Camera.position
				_eye.copy( _this.object.position ).sub( _this.target );

				//console.log('_this.object.position(2)='+_this.object.position.x + ', ' + this.object.position.y + ', ' + this.object.position.z);


				// 視点ベクトル
				//  ⇒注視点⇒視点ベクトルを正規化
				eyeDirection.copy( _eye ).normalize();


				// カメラ上方向ベクトル
				//  ⇒現在のカメラ上方向を正規化
				objectUpDirection.copy( _this.object.up ).normalize();


				// カメラ横方向ベクトル
				//  ⇒(視点ベクトル(1) × カメラ上方向ベクトル(1))
				objectSidewaysDirection.crossVectors( objectUpDirection, eyeDirection ).normalize();


				// カメラ上方向ベクトル(1)の長さをマウスY移動量で更新
				//  ⇒
				objectUpDirection.setLength( _moveCurr.y - _movePrev.y );


				// カメラ横方向ベクトル(1)の長さをマウスX移動量で更新
				//  ⇒
				objectSidewaysDirection.setLength( _moveCurr.x - _movePrev.x );


				// マウス移動ベクトル
				//  ⇒(カメラ上方向 + カメラ横方向)
				moveDirection.copy( objectUpDirection.add( objectSidewaysDirection ) );





				// 回転軸!!
				//  ⇒(マウス移動ベクトル × 視点ベクトル)
				//  ⇒クォータニオンを作るので正規化が必要
				axis.crossVectors( moveDirection, _eye ).normalize();


				// kawa add >>
				if(arrowHelper)
				{
					arrowHelper.setDirection( axis.clone().normalize() );
				}
				// kawa add <<


				angle *= _this.rotateSpeed;


				// 回転軸と回転角度からクォータニオンを作る
				quaternion.setFromAxisAngle( axis, angle );


				// 視点ベクトルを回転
				_eye.applyQuaternion( quaternion );


				// カメラ上方向ベクトルを回転
				//  ⇒これを回転させるとイマイチ直感的な回転にならない（メタセコのトラックボールも同じだった）
				//    ただし回転をやめるとカメラ向ベクトルとカメラ上方向ベクトルが同じになると挙動がおかしくなる。。。
				_this.object.up.applyQuaternion( quaternion );


				// つまり、注視点を軸に回している
				
				console.log("eye=("+_eye.x+", "+_eye.y+", "+_eye.z);
				console.log("up=("+_this.object.up.x+", "+_this.object.up.y+", "+_this.object.up.z);


				_lastAxis.copy( axis );
				_lastAngle = angle;

			} else if ( ! _this.staticMoving && _lastAngle ) {

				// ここの処理は、回転後の余韻を表現するため

				_lastAngle *= Math.sqrt( 1.0 - _this.dynamicDampingFactor );
				_eye.copy( _this.object.position ).sub( _this.target );
				quaternion.setFromAxisAngle( _lastAxis, _lastAngle );
				_eye.applyQuaternion( quaternion );
				_this.object.up.applyQuaternion( quaternion );

			}

			_movePrev.copy( _moveCurr );

		};

	}() );


	this.zoomCamera = function () {

		var factor;

		if ( _state === STATE.TOUCH_ZOOM_PAN ) {

			factor = _touchZoomDistanceStart / _touchZoomDistanceEnd;
			_touchZoomDistanceStart = _touchZoomDistanceEnd;
			_eye.multiplyScalar( factor );

		} else {

			factor = 1.0 + ( _zoomEnd.y - _zoomStart.y ) * _this.zoomSpeed;

			if ( factor !== 1.0 && factor > 0.0 ) {

				_eye.multiplyScalar( factor );

				if ( _this.staticMoving ) {

					_zoomStart.copy( _zoomEnd );

				} else {

					_zoomStart.y += ( _zoomEnd.y - _zoomStart.y ) * this.dynamicDampingFactor;

				}

			}

		}

	};

	// カメラのパン
	//
	// _panStart：ドラッグ開始座標(スクリーン座標を左上原点で正規化した座標。0～1)
	// _panEnd  ：ドラッグ終了座標(同上)
	this.panCamera = ( function() {

		var mouseChange = new THREE.Vector2(),
			objectUp = new THREE.Vector3(),
			pan = new THREE.Vector3();

		return function panCamera() {

			// 移動方向ベクトルを作成 ( _panEnd - _panStart )
			mouseChange.copy( _panEnd ).sub( _panStart );

			// 移動方向ベクトルの長さが0ではない?
			// lengthSqは、ノルムではなく、
			if ( mouseChange.lengthSq() ) {

				//console.log("_eye=(" + _eye.x + ", " + _eye.y + ", " + _eye.z + ")");// kawa add
				//console.log("_eye.length()=" + _eye.length() );// kawa add
				//console.log("_this.object.up=(" + _this.object.up.x + ", " + _this.object.up.y + ", " + _this.object.up.z + ")");// kawa add
				

				// _eyeは、カメラ位置(0,0,1000)
				// _eye.length()は、ノルムなので 1000
				mouseChange.multiplyScalar( _eye.length() * _this.panSpeed );

				// カメラの位置とカメラの上方向(0,1,0)の外積をとって、
				pan.copy( _eye ).cross( _this.object.up ).setLength( mouseChange.x );

				

				pan.add( objectUp.copy( _this.object.up ).setLength( mouseChange.y ) );

				_this.object.position.add( pan );
				_this.target.add( pan );

				if ( _this.staticMoving ) {

					_panStart.copy( _panEnd );

				} else {

					_panStart.add( mouseChange.subVectors( _panEnd, _panStart ).multiplyScalar( _this.dynamicDampingFactor ) );

				}

			}

		};

	}() );

	this.checkDistances = function () {

		if ( ! _this.noZoom || ! _this.noPan ) {

			if ( _eye.lengthSq() > _this.maxDistance * _this.maxDistance ) {

				_this.object.position.addVectors( _this.target, _eye.setLength( _this.maxDistance ) );
				_zoomStart.copy( _zoomEnd );

			}

			if ( _eye.lengthSq() < _this.minDistance * _this.minDistance ) {

				_this.object.position.addVectors( _this.target, _eye.setLength( _this.minDistance ) );
				_zoomStart.copy( _zoomEnd );

			}

		}

	};

	// kawa chg >>
	//this.update = function () {
	this.update = function (arrowHelper) {
	// kawa chg <<

		_eye.subVectors( _this.object.position, _this.target );

		if ( ! _this.noRotate ) {

			// kawa chg >>
			//_this.rotateCamera();
			_this.rotateCamera(arrowHelper);
			// kawa chg <<

		}

		if ( ! _this.noZoom ) {

			_this.zoomCamera();

		}

		if ( ! _this.noPan ) {

			_this.panCamera();

		}

		_this.object.position.addVectors( _this.target, _eye );

		_this.checkDistances();

		_this.object.lookAt( _this.target );

		if ( lastPosition.distanceToSquared( _this.object.position ) > EPS ) {

			_this.dispatchEvent( changeEvent );

			lastPosition.copy( _this.object.position );

		}

	};

	this.reset = function () {

		_state = STATE.NONE;
		_prevState = STATE.NONE;

		_this.target.copy( _this.target0 );
		_this.object.position.copy( _this.position0 );
		_this.object.up.copy( _this.up0 );

		_eye.subVectors( _this.object.position, _this.target );

		_this.object.lookAt( _this.target );

		_this.dispatchEvent( changeEvent );

		lastPosition.copy( _this.object.position );

	};

	// listeners

	function keydown( event ) {

		if ( _this.enabled === false ) return;

		_prevState = _state;

		if ( _state !== STATE.NONE ) {

			return;

		} else if ( event.keyCode === _this.keys[ STATE.ROTATE ] && ! _this.noRotate ) {

			_state = STATE.ROTATE;

		} else if ( event.keyCode === _this.keys[ STATE.ZOOM ] && ! _this.noZoom ) {

			_state = STATE.ZOOM;

		} else if ( event.keyCode === _this.keys[ STATE.PAN ] && ! _this.noPan ) {

			_state = STATE.PAN;

		}

	}

	function keyup( event ) {

		if ( _this.enabled === false ) return;

		_state = _prevState;

	}

	function mousedown( event ) {

		if ( _this.enabled === false ) return;

		if ( _state === STATE.NONE ) {

			_state = event.button;

		}

		if ( _state === STATE.ROTATE && ! _this.noRotate ) {

			_moveCurr.copy( getMouseOnCircle( event.pageX, event.pageY ) );
			_movePrev.copy( _moveCurr );

		} else if ( _state === STATE.ZOOM && ! _this.noZoom ) {

			_zoomStart.copy( getMouseOnScreen( event.pageX, event.pageY ) );
			_zoomEnd.copy( _zoomStart );

		} else if ( _state === STATE.PAN && ! _this.noPan ) {

			_panStart.copy( getMouseOnScreen( event.pageX, event.pageY ) );

			//console.log("_panStart=(" + _panStart.x + ", " + _panStart.y + ")");// kawa add

			_panEnd.copy( _panStart );

		}

		document.addEventListener( 'mousemove', mousemove, false );
		document.addEventListener( 'mouseup', mouseup, false );

		_this.dispatchEvent( startEvent );

	}

	function mousemove( event ) {

		if ( _this.enabled === false ) return;

		if ( _state === STATE.ROTATE && ! _this.noRotate ) {

			_movePrev.copy( _moveCurr );
			_moveCurr.copy( getMouseOnCircle( event.pageX, event.pageY ) );

		} else if ( _state === STATE.ZOOM && ! _this.noZoom ) {

			_zoomEnd.copy( getMouseOnScreen( event.pageX, event.pageY ) );

		} else if ( _state === STATE.PAN && ! _this.noPan ) {

			_panEnd.copy( getMouseOnScreen( event.pageX, event.pageY ) );

		}

	}

	function mouseup( event ) {

		if ( _this.enabled === false ) return;

		_state = STATE.NONE;

		document.removeEventListener( 'mousemove', mousemove );
		document.removeEventListener( 'mouseup', mouseup );
		_this.dispatchEvent( endEvent );

	}

	function mousewheel( event ) {

		if ( _this.enabled === false ) return;

		var delta = 0;

		if ( event.wheelDelta ) {

			// WebKit / Opera / Explorer 9

			delta = event.wheelDelta / 40;

		} else if ( event.detail ) {

			// Firefox

			delta = - event.detail / 3;

		}

		_zoomStart.y += delta * 0.01;
		_this.dispatchEvent( startEvent );
		_this.dispatchEvent( endEvent );

	}

	function touchstart( event ) {

		if ( _this.enabled === false ) return;

		switch ( event.touches.length ) {

			case 1:
				_state = STATE.TOUCH_ROTATE;
				_moveCurr.copy( getMouseOnCircle( event.touches[ 0 ].pageX, event.touches[ 0 ].pageY ) );
				_movePrev.copy( _moveCurr );
				break;

			default: // 2 or more
				_state = STATE.TOUCH_ZOOM_PAN;
				var dx = event.touches[ 0 ].pageX - event.touches[ 1 ].pageX;
				var dy = event.touches[ 0 ].pageY - event.touches[ 1 ].pageY;
				_touchZoomDistanceEnd = _touchZoomDistanceStart = Math.sqrt( dx * dx + dy * dy );

				var x = ( event.touches[ 0 ].pageX + event.touches[ 1 ].pageX ) / 2;
				var y = ( event.touches[ 0 ].pageY + event.touches[ 1 ].pageY ) / 2;
				_panStart.copy( getMouseOnScreen( x, y ) );
				_panEnd.copy( _panStart );
				break;

		}

		_this.dispatchEvent( startEvent );

	}

	function touchmove( event ) {

		if ( _this.enabled === false ) return;

		switch ( event.touches.length ) {

			case 1:
				_movePrev.copy( _moveCurr );
				_moveCurr.copy( getMouseOnCircle( event.touches[ 0 ].pageX, event.touches[ 0 ].pageY ) );
				break;

			default: // 2 or more
				var dx = event.touches[ 0 ].pageX - event.touches[ 1 ].pageX;
				var dy = event.touches[ 0 ].pageY - event.touches[ 1 ].pageY;
				_touchZoomDistanceEnd = Math.sqrt( dx * dx + dy * dy );

				var x = ( event.touches[ 0 ].pageX + event.touches[ 1 ].pageX ) / 2;
				var y = ( event.touches[ 0 ].pageY + event.touches[ 1 ].pageY ) / 2;
				_panEnd.copy( getMouseOnScreen( x, y ) );
				break;

		}

	}

	function touchend( event ) {

		if ( _this.enabled === false ) return;

		switch ( event.touches.length ) {

			case 0:
				_state = STATE.NONE;
				break;

			case 1:
				_state = STATE.TOUCH_ROTATE;
				_moveCurr.copy( getMouseOnCircle( event.touches[ 0 ].pageX, event.touches[ 0 ].pageY ) );
				_movePrev.copy( _moveCurr );
				break;

		}

		_this.dispatchEvent( endEvent );

	}

	function contextmenu( event ) {

		event.preventDefault();

	}

	this.dispose = function() {

		this.domElement.removeEventListener( 'contextmenu', contextmenu, false );
		this.domElement.removeEventListener( 'mousedown', mousedown, false );
		this.domElement.removeEventListener( 'mousewheel', mousewheel, false );
		this.domElement.removeEventListener( 'MozMousePixelScroll', mousewheel, false ); // firefox

		this.domElement.removeEventListener( 'touchstart', touchstart, false );
		this.domElement.removeEventListener( 'touchend', touchend, false );
		this.domElement.removeEventListener( 'touchmove', touchmove, false );

		document.removeEventListener( 'mousemove', mousemove, false );
		document.removeEventListener( 'mouseup', mouseup, false );

		window.removeEventListener( 'keydown', keydown, false );
		window.removeEventListener( 'keyup', keyup, false );

	};

	this.domElement.addEventListener( 'contextmenu', contextmenu, false );
	this.domElement.addEventListener( 'mousedown', mousedown, false );
	this.domElement.addEventListener( 'mousewheel', mousewheel, false );
	this.domElement.addEventListener( 'MozMousePixelScroll', mousewheel, false ); // firefox

	this.domElement.addEventListener( 'touchstart', touchstart, false );
	this.domElement.addEventListener( 'touchend', touchend, false );
	this.domElement.addEventListener( 'touchmove', touchmove, false );

	window.addEventListener( 'keydown', keydown, false );
	window.addEventListener( 'keyup', keyup, false );

	this.handleResize();

	// force an update at start
	this.update();

};

THREE.TrackballControls.prototype = Object.create( THREE.EventDispatcher.prototype );
THREE.TrackballControls.prototype.constructor = THREE.TrackballControls;
