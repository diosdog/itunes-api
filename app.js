var fs = require('fs')
var path = require('path')
var util = require('util')
var express = require('express')
var morgan = require('morgan')
var bodyParser = require('body-parser')
var parameterize = require('parameterize');
const shell = require('node-powershell');
var winax = require('winax');
var tmp = require('tmp');

var app = express()
app.use(bodyParser.urlencoded({ extended: false }))
app.use(express.static(path.join(__dirname, 'public')));

var logFormat = "'[:date[iso]] - :remote-addr - :method :url :status :response-time ms - :res[content-length]b'"
app.use(morgan(logFormat))

var tmpDir = null;


function win_getCurrentState(){
	itunes = new ActiveXObject('iTunes.Application')

	currentState = {};

	currentState['player_state'] = itunes.PlayerState ? 'playing' : 'paused' //TODO: convert to expected format to match mac version instead of 0/1 for paused/playing

	currentTrack = itunes.CurrentTrack
	currentPlaylist = itunes.CurrentPlaylist

	currentState['id'] = currentTrack.trackID.toString()
	currentState['name'] = currentTrack.Name
	currentState['artist'] = currentTrack.Artist
    currentState['album'] = currentTrack.Album
    currentState['playlist'] = currentPlaylist.Favorites
    currentState['volume'] = itunes.SoundVolume
    currentState['muted'] = itunes.Mute
    currentState['repeat'] = 0 //TODO: figure out if this is accessible in COM object
    currentState['shuffle'] = 0 //TODO: figure out if this is accessible in COM object

	if (currentTrack.Year) {
		currentState['album'] += " (" + currentTrack.Year + ")";
	}

	winax.release(itunes)

	return currentState
}

function getCurrentState(){
  itunes = Application('iTunes');
  playerState = itunes.playerState();
  currentState = {};

  currentState['player_state'] = playerState;

  if (playerState != "stopped") {
    currentTrack = itunes.currentTrack;
    currentPlaylist = itunes.currentPlaylist;

    currentState['id'] = currentTrack.persistentID();
    currentState['name'] = currentTrack.name();
    currentState['artist'] = currentTrack.artist();
    currentState['album'] = currentTrack.album();
    currentState['playlist'] = currentPlaylist.name();
    currentState['volume'] = itunes.soundVolume();
    currentState['muted'] = itunes.mute();
    currentState['repeat'] = itunes.songRepeat();
    currentState['shuffle'] = itunes.shuffleEnabled() && itunes.shuffleMode();

    if (currentTrack.year()) {
      currentState['album'] += " (" + currentTrack.year() + ")";
    }
  }

  return currentState;
}

function sendResponse(error, res){
  if (error) {
    console.log(error)
    res.sendStatus(500)
  }else{
    //osa(getCurrentState, function (error, state) {
	try {
		state = win_getCurrentState()
		res.json(state)
	} catch(error) {
		console.log(error)
		res.sendStatus(500)
	}
  }
}

function win_playPlaylist(nameOrId) {
	//TODO: find if this is accessible in COM object
	return false;
}

function playPlaylist(nameOrId){
  itunes = Application('iTunes');

  if ((nameOrId - 0) == nameOrId && ('' + nameOrId).trim().length > 0) {
    id = parseInt(nameOrId);
    itunes.playlists.byId(id).play();
  }else{
    itunes.playlists.byName(nameOrId).play();
  }

  return true;
}

function win_setVolume(level){
	itunes = new ActiveXObject('iTunes.Application');
	var ret = true;
	if (level) {
		itunes.SoundVolume = parseInt(level);
	} else {
		ret = false;
	}
	winax.release(itunes)
	return ret;
}

function setVolume(level){
  itunes = Application('iTunes');

  if (level) {
    itunes.soundVolume = parseInt(level);
    return true;
  }else {
    return false;
  }
}

function win_setMuted(muted){
	itunes = new ActiveXObject('iTunes.Application');
	var ret = true;
	if (muted) {
		itunes.Mute = muted;
	} else {
		ret = false;
	}
	winax.release(itunes)
	return ret;
}

function setMuted(muted){
  itunes = Application('iTunes');

  if (muted) {
    itunes.mute = muted;
    return true;
  }else{
    return false;
  }
}

function win_setShuffle(mode) {
	//TODO: figure out if com object supports this
	return false;
}

function setShuffle(mode){
  itunes = Application('iTunes');

  if (!mode) {
    mode = "songs"
  }

  if (mode == "false" || mode == "off") {
    itunes.shuffleEnabled = false;
    return false;
  }else{
    itunes.shuffleEnabled = true;
    itunes.shuffleMode = mode;
    return true;
  }
}

function win_setRepeat(mode){
	//TODO: figure out if com object supports this
	return false;
}

function setRepeat(mode){
  itunes = Application('iTunes');

  if (!mode) {
    mode = "all"
  }

  if (mode == "false" || mode == "off") {
    itunes.songRepeat = false;
    return false;
  }else{
    itunes.songRepeat = mode;
    return true;
  }
}

function win_getPlaylistsFromItunes(){

	itunes = new ActiveXObject('iTunes.Application');
	playlists = itunes.GetITObjectByID(itunes.LibrarySource.Index,0,0,0).Playlists

    playlistNames = [];

    for (var i = 0; i < playlists.length; i++) {
      playlist = playlists[i];

      data = {};
      data['id'] = playlist.Index;
      data['name'] = playlist.Name;
      data['loved'] = false; //TODO
      data['duration_in_seconds'] = playlist.Duration;
      data['time'] = playlist.Time;
      playlistNames.push(data);
    }

	winax.release(itunes);

    return playlistNames;
}

function getPlaylistsFromItunes(){
  itunes = Application('iTunes');
  playlists = itunes.playlists();

  playlistNames = [];

  for (var i = 0; i < playlists.length; i++) {
    playlist = playlists[i];

    data = {};
    data['id'] = playlist.id();
    data['name'] = playlist.name();
    data['loved'] = playlist.loved();
    data['duration_in_seconds'] = playlist.duration();
    data['time'] = playlist.time();
    playlistNames.push(data);
  }

  return playlistNames;
}

function getPlaylists(callback){
	data = win_getPlaylistsFromItunes();
	//(not sure if this is necessary for win version)
	for (var i = 0; i < data.length; i++) {
	  data[i]['id'] = parameterize(data[i]['name'])
	}
}

app.get('/_ping', function(req, res){
  res.send('OK');
})

app.get('/', function(req, res){
  res.sendfile('index.html');
})


app.put('/play', function(req, res){
	try {
		itunes = new ActiveXObject('iTunes.Application');
		itunes.Play();
		winax.release(itunes);
		sendResponse(null,res);
	} catch(error) {
		sendResponse(error, res);
	}
})

app.put('/pause', function(req, res){
	try {
  	  itunes = new ActiveXObject('iTunes.Application');
  	  itunes.Pause();
  	  sendResponse(null,res);
    } catch(error) {
  	  sendResponse(error, res);
    }
})

app.put('/playpause', function(req, res){
	try {
  	  itunes = new ActiveXObject('iTunes.Application');
  	  itunes.PlayPause();
  	  winax.release(itunes);
  	  sendResponse(null,res);
    } catch(error) {
  	  sendResponse(error, res);
    }
})

app.put('/stop', function(req, res){
	try {
  	  itunes = new ActiveXObject('iTunes.Application');
  	  itunes.Stop()
  	  winax.release(itunes);
  	  sendResponse(null,res);
    } catch(error) {
  	  sendResponse(error, res);
    }
})

app.put('/previous', function(req, res){
	try {
  	  itunes = new ActiveXObject('iTunes.Application');
  	  itunes.PreviousTrack()
  	  winax.release(itunes);
  	  sendResponse(null,res);
    } catch(error) {
  	  sendResponse(error, res);
    }
})

app.put('/next', function(req, res){
	try {
  	  itunes = new ActiveXObject('iTunes.Application');
  	  itunes.NextTrack()
  	  winax.release(itunes);
  	  sendResponse(null,res);
    } catch(error) {
  	  sendResponse(error, res);
    }
})

app.put('/volume', function(req, res){
	try {
		win_setVolume(req.body.level);
		sendResponse(null,res);
	} catch(error) {
		sendResponse(error, res);
	}
})

app.put('/mute', function(req, res){
	try {
		win_setMuted(req.body.muted);
		sendResponse(null,res);
	} catch(error) {
		sendResponse(error, res);
	}
})

app.put('/shuffle', function(req, res){
	try {
		win_setShuffle(req.body.mode);
		sendResponse(null,res);
	} catch(error) {
		sendResponse(error, res);
	}
})

app.put('/repeat', function(req, res){
	try {
		win_setRepeat(req.body.mode);
		sendResponse(null,res);
	} catch(error) {
		sendResponse(error, res);
	}
})

app.get('/now_playing', function(req, res){
  error = null
  sendResponse(error, res)
})

app.get('/artwork', function(req, res){
	try {
		var tmpFile = tmp.tmpNameSync() + ".png"; //TODO: handle other formats rather than just assuming png
		itunes = new ActiveXObject('iTunes.Application');
		itunes.CurrentTrack.Artwork(1).SaveArtworkToFile(tmpFile);
		winax.release(itunes);

		res.type = 'image/png';
		res.sendFile(tmpFile)
	} catch(error) {
		sendResponse(error, res)
	}
})

app.get('/playlists', function (req, res) {
	sendResponse("Not supported",res);
	return;
	try {
		data = getPlaylists();
		res.json({playlists: data});
	} catch(error) {
		sendResponse(error,res);
	}
})

app.put('/playlists/:id/play', function (req, res) {
	sendResponse("Not supported",res); //TODO: add support for win version
})

app.get('/airplay_devices', function(req, res){
  //sendResponse("Not supported",res); //TODO: add support for win version
  res.json({'airplay_devices': {}})
})

app.get('/airplay_devices/:id', function(req, res){
  sendResponse("Not supported",res); //TODO: add support for win version
})

app.put('/airplay_devices/:id/on', function (req, res) {
  sendResponse("Not supported",res); //TODO: add support for win version
})

app.put('/airplay_devices/:id/off', function (req, res) {
  sendResponse("Not supported",res); //TODO: add support for win version
})

app.put('/airplay_devices/:id/volume', function (req, res) {
  sendResponse("Not supported",res); //TODO: add support for win version
})

app.listen(process.env.PORT || 8181);
