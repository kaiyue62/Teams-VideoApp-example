microsoftTeams.initialize(() => {}, ["https://localhost:9000", "https://lubobill1990.github.io"]);

// This is the effect for processing
let appliedEffect = {
  pixelValue: 100,
  proportion: 2,
};

// This is the effect linked with UI
let uiSelectedEffect = {};

let errorType={
  none:0,
  slow:1,
  frozen:2
}

let error=errorType.none;

let errorOccurs = false;

function videoFrameErrorHandler(videoFrame, notifyVideoProcessed, notifyError){
  let timeout=0;
  if(uiSelectedEffect.timeout>0){
    timeout=uiSelectedEffect.timeout
  }else if(error===errorType.slow){
    timeout=200;
  }else if(error===errorType.frozen){
    timeout=2000;
  }else if(error===errorType.none){
    timeout=0;
  }

  setTimeout(() => {
    videoFrameHandler(videoFrame, notifyVideoProcessed, notifyError);
  }, timeout);
}

//Sample video effect
function videoFrameHandler(videoFrame, notifyVideoProcessed, notifyError) {
  const maxLen =
    (videoFrame.height * videoFrame.width) /
    Math.max(1, appliedEffect.proportion);

  for (let i = 0; i < maxLen; i += 4) {
    //smaple effect just change the value to 100, which effect some pixel value of video frame
    videoFrame.data[i + 1] = appliedEffect.pixelValue;
  }

  //send notification the effect processing is finshed.
  notifyVideoProcessed();

  //send error to Teams
  if (errorOccurs) {
    notifyError("some error message");
  }
}

function effectParameterChanged(effectName) {
  console.log(effectName);
  if (effectName === undefined) {
    // If effectName is undefined, then apply the effect selected in the UI
    appliedEffect = {
      ...appliedEffect,
      ...uiSelectedEffect,
    };
  } else {
    // if effectName is string sent from Teams client, the apply the effectName
    try {
      appliedEffect = {
        ...appliedEffect,
        ...JSON.parse(effectName),
      };
    } catch (e) {}
  }
}

microsoftTeams.video.registerForVideoEffect(effectParameterChanged);
microsoftTeams.video.registerForVideoFrame(videoFrameErrorHandler, {
  format: "NV12",
});

// any changes to the UI should notify Teams client.
document.getElementById("enable_check").addEventListener("change", function () {
  if (this.checked) {
    microsoftTeams.video.notifySelectedVideoEffectChanged("EffectChanged");
  } else {
    microsoftTeams.video.notifySelectedVideoEffectChanged("EffectDisabled");
  }
});
document.getElementById("proportion").addEventListener("change", function () {
  uiSelectedEffect.proportion = this.value;
  microsoftTeams.video.notifySelectedVideoEffectChanged("EffectChanged");
});
document.getElementById("pixel_value").addEventListener("change", function () {
  uiSelectedEffect.pixelValue = this.value;
  microsoftTeams.video.notifySelectedVideoEffectChanged("EffectChanged");
});
document.getElementById('slow-btn').addEventListener('click',function(){
  error=errorType.slow;
});
document.getElementById('frozen-btn').addEventListener('click',function(){
  error=errorType.frozen;
});
document.getElementById('reset-btn').addEventListener('click',function(){
  error=errorType.none;
});
document.getElementById("timeout_val").addEventListener("change", function () {
  uiSelectedEffect.timeout = this.value
});
microsoftTeams.appInitialization.notifySuccess();
