function catLoader() {
  return `
        <!-- From Uiverse.io by vinodjangid07 --> 
        <div class="loader">
        <div class="wrapper">
            <div class="catContainer">
            <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 733 673"
                class="catbody"
            >
                <path
                fill="#212121"
                d="M111.002 139.5C270.502 -24.5001 471.503 2.4997 621.002 139.5C770.501 276.5 768.504 627.5 621.002 649.5C473.5 671.5 246 687.5 111.002 649.5C-23.9964 611.5 -48.4982 303.5 111.002 139.5Z"
                ></path>
                <path fill="#212121" d="M184 9L270.603 159H97.3975L184 9Z"></path>
                <path fill="#212121" d="M541 0L627.603 150H454.397L541 0Z"></path>
            </svg>
            <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 158 564"
                class="tail"
            >
                <path
                fill="#191919"
                d="M5.97602 76.066C-11.1099 41.6747 12.9018 0 51.3036 0V0C71.5336 0 89.8636 12.2558 97.2565 31.0866C173.697 225.792 180.478 345.852 97.0691 536.666C89.7636 553.378 73.0672 564 54.8273 564V564C16.9427 564 -5.4224 521.149 13.0712 488.085C90.2225 350.15 87.9612 241.089 5.97602 76.066Z"
                ></path>
            </svg>
            <div class="text">
                <span class="bigzzz">Z</span>
                <span class="zzz">Z</span>
            </div>
            </div>
            <div class="wallContainer">
            <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 500 126"
                class="wall"
            >
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="3"
                x2="450"
                y1="3"
                x1="50"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="85"
                x2="400"
                y1="85"
                x1="100"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="122"
                x2="375"
                y1="122"
                x1="125"
                ></line>
                <line stroke-width="6" stroke="#7C7C7C" y2="43" x2="500" y1="43"></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="1.99391"
                x2="115.5"
                y1="43.0061"
                x1="115.5"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="2.00002"
                x2="189"
                y1="43.0122"
                x1="189"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="2.00612"
                x2="262.5"
                y1="43.0183"
                x1="262.5"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="2.01222"
                x2="336"
                y1="43.0244"
                x1="336"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="2.01833"
                x2="409.5"
                y1="43.0305"
                x1="409.5"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="43"
                x2="153"
                y1="84.0122"
                x1="153"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="43"
                x2="228"
                y1="84.0122"
                x1="228"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="43"
                x2="303"
                y1="84.0122"
                x1="303"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="43"
                x2="378"
                y1="84.0122"
                x1="378"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="84"
                x2="192"
                y1="125.012"
                x1="192"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="84"
                x2="267"
                y1="125.012"
                x1="267"
                ></line>
                <line
                stroke-width="6"
                stroke="#7C7C7C"
                y2="84"
                x2="342"
                y1="125.012"
                x1="342"
                ></line>
            </svg>
            </div>
        </div>
        </div>
        <style>
        /* From Uiverse.io by vinodjangid07 */ 
        .loader {
        width: 100%;
        height: fit-content;
        display: flex;
        align-items: center;
        justify-content: center;
        }
        .wrapper {
        width: fit-content;
        height: fit-content;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        }
        .catContainer {
        width: 100%;
        height: fit-content;
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
        }
        .catbody {
        width: 80px;
        }
        .tail {
        position: absolute;
        width: 17px;
        top: 50%;
        animation: tail 0.5s ease-in infinite alternate-reverse;
        transform-origin: top;
        }
        @keyframes tail {
        0% {
            transform: rotateZ(60deg);
        }
        50% {
            transform: rotateZ(0deg);
        }
        100% {
            transform: rotateZ(-20deg);
        }
        }
        .wall {
        width: 300px;
        }
        .text {
        display: flex;
        flex-direction: column;
        width: 50px;
        position: absolute;
        margin: 0px 0px 100px 120px;
        }
        .zzz {
        color: black;
        font-weight: 700;
        font-size: 15px;
        animation: zzz 2s linear infinite;
        }
        .bigzzz {
        color: black;
        font-weight: 700;
        font-size: 25px;
        margin-left: 10px;
        animation: zzz 2.3s linear infinite;
        }
        @keyframes zzz {
        0% {
            color: transparent;
        }
        50% {
            color: black;
        }
        100% {
            color: transparent;
        }
        }
        </style>

    `;
}

function getMario() {
  return `
  <!-- From Uiverse.io by https://uiverse.io/vikramsinghnegi/dangerous-fireant-25 --> 
<div class="loader-mario"></div>
<style>
  .loader-mario {
    width: fit-content;
    font-size: 17px;
    font-family: monospace;
    line-height: 1.4;
    font-weight: bold;
    padding: 30px 2px 50px;
    background: linear-gradient(#000 0 0) 0 0/100% 100% content-box padding-box
      no-repeat;
    position: relative;
    overflow: hidden;
    animation: l10-0 2s infinite cubic-bezier(1, 175, 0.5, 175);
  }
  .loader-mario::before {
    content: "Loading";
    display: inline-block;
    animation: l10-2 2s infinite;
  }
  .loader-mario::after {
    content: "";
    position: absolute;
    width: 34px;
    height: 28px;
    top: 110%;
    left: calc(50% - 16px);
    background: linear-gradient(
          90deg,
          #0000 12px,
          #f92033 0 22px,
          #0000 0 26px,
          #fdc98d 0 32px,
          #0000
        )
        bottom 26px left 50%,
      linear-gradient(90deg, #0000 10px, #f92033 0 28px, #fdc98d 0 32px, #0000 0)
        bottom 24px left 50%,
      linear-gradient(
          90deg,
          #0000 10px,
          #643700 0 16px,
          #fdc98d 0 20px,
          #000 0 22px,
          #fdc98d 0 24px,
          #000 0 26px,
          #f92033 0 32px,
          #0000 0
        )
        bottom 22px left 50%,
      linear-gradient(
          90deg,
          #0000 8px,
          #643700 0 10px,
          #fdc98d 0 12px,
          #643700 0 14px,
          #fdc98d 0 20px,
          #000 0 22px,
          #fdc98d 0 28px,
          #f92033 0 32px,
          #0000 0
        )
        bottom 20px left 50%,
      linear-gradient(
          90deg,
          #0000 8px,
          #643700 0 10px,
          #fdc98d 0 12px,
          #643700 0 16px,
          #fdc98d 0 22px,
          #000 0 24px,
          #fdc98d 0 30px,
          #f92033 0 32px,
          #0000 0
        )
        bottom 18px left 50%,
      linear-gradient(
          90deg,
          #0000 8px,
          #643700 0 12px,
          #fdc98d 0 20px,
          #000 0 28px,
          #f92033 0 30px,
          #0000 0
        )
        bottom 16px left 50%,
      linear-gradient(90deg, #0000 12px, #fdc98d 0 26px, #f92033 0 30px, #0000 0)
        bottom 14px left 50%,
      linear-gradient(
          90deg,
          #fdc98d 6px,
          #f92033 0 14px,
          #222a87 0 16px,
          #f92033 0 22px,
          #222a87 0 24px,
          #f92033 0 28px,
          #0000 0 32px,
          #643700 0
        )
        bottom 12px left 50%,
      linear-gradient(
          90deg,
          #fdc98d 6px,
          #f92033 0 16px,
          #222a87 0 18px,
          #f92033 0 24px,
          #f92033 0 26px,
          #0000 0 30px,
          #643700 0
        )
        bottom 10px left 50%,
      linear-gradient(
          90deg,
          #0000 10px,
          #f92033 0 16px,
          #222a87 0 24px,
          #feee49 0 26px,
          #222a87 0 30px,
          #643700 0
        )
        bottom 8px left 50%,
      linear-gradient(
          90deg,
          #0000 12px,
          #222a87 0 18px,
          #feee49 0 20px,
          #222a87 0 30px,
          #643700 0
        )
        bottom 6px left 50%,
      linear-gradient(90deg, #0000 8px, #643700 0 12px, #222a87 0 30px, #643700 0)
        bottom 4px left 50%,
      linear-gradient(90deg, #0000 6px, #643700 0 14px, #222a87 0 26px, #0000 0)
        bottom 2px left 50%,
      linear-gradient(90deg, #0000 6px, #643700 0 10px, #0000 0) bottom 0px left
        50%;
    background-size: 34px 2px;
    background-repeat: no-repeat;
    animation: inherit;
    animation-name: l10-1;
  }
  @keyframes l10-0 {
    0%,
    30% {
      background-position: 0 0px;
    }
    50%,
    100% {
      background-position: 0 -0.1px;
    }
  }
  @keyframes l10-1 {
    50%,
    100% {
      top: 109.5%;
    }
  }
  @keyframes l10-2 {
    0%,
    30% {
      transform: translateY(0);
    }
    80%,
    100% {
      transform: translateY(-260%);
    }
  }
</style>

  `;
}

function getLoader() {

  const loader1 = `
<!-- From Uiverse.io by https://uiverse.io/dilmurod_3869/new-penguin-19 --> 
<div class="typewriter-alt">
  <div class="slide"><i></i></div>
  <div class="paper"></div>
  <div class="keyboard"></div>
</div>

<style>
/* From Uiverse.io by https://uiverse.io/dilmurod_3869/new-penguin-19 */ 
.typewriter-alt {
  --green: #5c86ff;
  --green-dark: #1e0325;
  --key: #fff;
  --paper: #eef0fd;
  --text: #d3d4ec;
  --tool: #fbc56c;  --duration: 2.5s;
  position: relative;
  animation: bounce-alt var(--duration) ease-in-out infinite;
}

.typewriter-alt .slide {
  width: 100px;
  height: 18px;
  border-radius: 4px;
  margin-left: 10px;
  transform: translateX(10px);
  background: linear-gradient(var(--green), var(--green-dark));
  animation: slide-alt var(--duration) ease infinite;
}

.typewriter-alt .slide:before,
.typewriter-alt .slide:after,
.typewriter-alt .slide i:before {
  content: "";
  position: absolute;
  background: var(--tool);
}

.typewriter-alt .slide:before {
  width: 3px;
  height: 10px;
  top: 4px;
  left: 100%;
}

.typewriter-alt .slide:after {
  left: 102px;
  top: 2px;
  height: 12px;
  width: 5px;
  border-radius: 2px;
}

.typewriter-alt .slide i {
  display: block;
  position: absolute;
  right: 100%;
  width: 5px;
  height: 5px;
  top: 3px;
  background: var(--tool);
}

.typewriter-alt .slide i:before {
  right: 100%;
  top: -3px;
  width: 3px;
  border-radius: 1px;
  height: 12px;
}

.typewriter-alt .paper {
  position: absolute;
  left: 20px;
  top: -30px;
  width: 45px;
  height: 50px;
  border-radius: 6px;
  background: var(--paper);
  transform: translateY(50px);
  animation: paper-alt var(--duration) linear infinite;
}

.typewriter-alt .paper:before {
  content: "";
  position: absolute;
  left: 5px;
  right: 5px;
  top: 8px;
  border-radius: 1px;
  height: 3px;
  transform: scaleY(0.9);
  background: var(--text);
  box-shadow:
    0 10px 0 var(--text),
    0 20px 0 var(--text),
    0 30px 0 var(--text);
}

.typewriter-alt .keyboard {
  width: 130px;
  height: 60px;
  margin-top: -8px;
  z-index: 1;
  position: relative;
}

.typewriter-alt .keyboard:before,
.typewriter-alt .keyboard:after {
  content: "";
  position: absolute;
}

.typewriter-alt .keyboard:before {
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  border-radius: 8px;
  background: linear-gradient(135deg, var(--green), var(--green-dark));
  transform: perspective(12px) rotateX(3deg);
  transform-origin: 50% 100%;
}

.typewriter-alt .keyboard:after {
  left: 3px;
  top: 28px;
  width: 10px;
  height: 3px;
  border-radius: 1px;
  box-shadow:
    16px 0 0 var(--key),
    32px 0 0 var(--key),
    48px 0 0 var(--key),
    64px 0 0 var(--key),
    80px 0 0 var(--key),
    96px 0 0 var(--key),
    24px 8px 0 var(--key),
    40px 8px 0 var(--key),
    56px 8px 0 var(--key),
    64px 8px 0 var(--key),
    72px 8px 0 var(--key),
    88px 8px 0 var(--key);
  animation: keyboard-alt var(--duration) linear infinite;
}

@keyframes bounce-alt {
  0%,
  80%,
  100% {
    transform: translateY(0);
  }
  40% {
    transform: translateY(-6px);
  }
  60% {
    transform: translateY(3px);
  }
}

@keyframes slide-alt {
  0%,
  100% {
    transform: translateX(10px);
  }
  25% {
    transform: translateX(4px);
  }
  50% {
    transform: translateX(-2px);
  }
  75% {
    transform: translateX(-8px);
  }
}

@keyframes paper-alt {
  0%,
  100% {
    transform: translateY(50px);
  }
  25% {
    transform: translateY(40px);
  }
  50% {
    transform: translateY(25px);
  }
  75% {
    transform: translateY(10px);
  }
}

@keyframes keyboard-alt {
  0%,
  20%,
  40%,
  60%,
  80% {
    box-shadow:
      16px 0 0 var(--key),
      32px 0 0 var(--key),
      48px 0 0 var(--key),
      64px 0 0 var(--key),
      80px 0 0 var(--key),
      96px 0 0 var(--key),
      24px 8px 0 var(--key),
      40px 8px 0 var(--key),
      56px 8px 0 var(--key),
      64px 8px 0 var(--key),
      72px 8px 0 var(--key),
      88px 8px 0 var(--key);
  }
  10% {
    box-shadow:
      16px 2px 0 var(--key),
      32px 0 0 var(--key),
      48px 0 0 var(--key),
      64px 0 0 var(--key),
      80px 0 0 var(--key),
      96px 0 0 var(--key),
      24px 8px 0 var(--key),
      40px 8px 0 var(--key),
      56px 8px 0 var(--key),
      64px 8px 0 var(--key),
      72px 8px 0 var(--key),
      88px 8px 0 var(--key);
  }
  30% {
    box-shadow:
      16px 0 0 var(--key),
      32px 0 0 var(--key),
      48px 0 0 var(--key),
      64px 2px 0 var(--key),
      80px 0 0 var(--key),
      96px 0 0 var(--key),
      24px 8px 0 var(--key),
      40px 8px 0 var(--key),
      56px 8px 0 var(--key),
      64px 8px 0 var(--key),
      72px 8px 0 var(--key),
      88px 8px 0 var(--key);
  }
  50% {
    box-shadow:
      16px 0 0 var(--key),
      32px 0 0 var(--key),
      48px 0 0 var(--key),
      64px 0 0 var(--key),
      80px 0 0 var(--key),
      96px 0 0 var(--key),
      24px 10px 0 var(--key),
      40px 8px 0 var(--key),
      56px 8px 0 var(--key),
      64px 8px 0 var(--key),
      72px 8px 0 var(--key),
      88px 8px 0 var(--key);
  }
  70% {
    box-shadow:
      16px 0 0 var(--key),
      32px 0 0 var(--key),
      48px 0 0 var(--key),
      64px 0 0 var(--key),
      80px 2px 0 var(--key),
      96px 0 0 var(--key),
      24px 8px 0 var(--key),
      40px 8px 0 var(--key),
      56px 8px 0 var(--key),
      64px 8px 0 var(--key),
      72px 8px 0 var(--key),
      88px 8px 0 var(--key);
  }
}
</style>
`;

  const loader2 = `
<!-- From Uiverse.io by https://uiverse.io/vajion_7943/selfish-hound-5 -->
<svg
  id="svg_svg"
  xmlns="http://www.w3.org/2000/svg"
  fill="none"
  viewBox="0 0 477 578"
  height="434"
  width="358"
>
  <g filter="url(#filter0_i_163_1030)">
    <path
      fill="transparent"
      d="M235.036 304.223C236.949 303.118 240.051 303.118 241.964 304.223L470.072 435.921C473.898 438.13 473.898 441.712 470.072 443.921L247.16 572.619C242.377 575.38 234.623 575.38 229.84 572.619L6.92817 443.921C3.10183 441.712 3.10184 438.13 6.92817 435.921L235.036 304.223Z"
    ></path>
  </g>
  <path
    stroke="white"
    d="M235.469 304.473C237.143 303.506 239.857 303.506 241.531 304.473L469.639 436.171C473.226 438.242 473.226 441.6 469.639 443.671L246.727 572.369C242.183 574.992 234.817 574.992 230.273 572.369L7.36118 443.671C3.77399 441.6 3.774 438.242 7.36119 436.171L235.469 304.473Z"
  ></path>
  <path
    stroke="white"
    fill="#C3CADC"
    d="M234.722 321.071C236.396 320.105 239.111 320.105 240.785 321.071L439.477 435.786C443.064 437.857 443.064 441.215 439.477 443.286L240.785 558.001C239.111 558.967 236.396 558.967 234.722 558.001L36.0304 443.286C32.4432 441.215 32.4432 437.857 36.0304 435.786L234.722 321.071Z"
  ></path>
  <path
    fill="#4054B2"
    d="M234.521 366.089C236.434 364.985 239.536 364.985 241.449 366.089L406.439 461.346L241.247 556.72C239.333 557.825 236.231 557.825 234.318 556.72L69.3281 461.463L234.521 366.089Z"
  ></path>
  <path
    fill="#30439B"
    d="M237.985 364.089L237.984 556.972C236.144 556.941 235.082 556.717 233.13 556.043L69.3283 461.463L237.985 364.089Z"
  ></path>
  <path
    fill="url(#paint0_linear_163_1030)"
    d="M36.2146 117.174L237.658 0.435217V368.615C236.541 368.598 235.686 368.977 233.885 370.124L73.1836 463.678L39.2096 444.075C37.0838 442.229 36.285 440.981 36.2146 438.027V117.174Z"
    id="layer_pared"
  ></path>
  <path
    fill="url(#paint1_linear_163_1030)"
    d="M439.1 116.303L237.657 0.435568V368.616C238.971 368.585 239.822 369.013 241.43 370.135L403.64 462.925L436.128 444.089C437.832 442.715 438.975 441.147 439.1 439.536V116.303Z"
    id="layer_pared"
  ></path>
  <path
    fill="#27C6FD"
    d="M64.5447 181.554H67.5626V186.835L64.5447 188.344V181.554Z"
    id="float_server"
  ></path>
  <path
    fill="#138EB9"
    d="M88.3522 374.347L232.415 457.522C234.202 458.405 234.866 458.629 236.335 458.71V468.291C235.356 468.291 234.086 468.212 232.415 467.275L88.3522 384.1C86.3339 382.882 85.496 382.098 85.4707 380.198V370.428L88.3522 374.347Z"
    id="float_server"
  ></path>
  <path
    fill="#138EB9"
    d="M384.318 374.445L240.254 457.62C238.914 458.385 238.295 458.629 236.335 458.71V468.291C237.315 468.291 238.704 468.211 240.236 467.274L384.318 384.198C386.457 383.091 387.151 382.244 387.258 380.228V370.917C386.768 372.387 386.21 373.295 384.318 374.445Z"
    id="float_server"
  ></path>
  <path
    stroke="url(#paint3_linear_163_1030)"
    fill="url(#paint2_linear_163_1030)"
    d="M240.452 226.082L408.617 323.172C412.703 325.531 412.703 329.355 408.617 331.713L240.452 428.803C238.545 429.904 235.455 429.904 233.548 428.803L65.3832 331.713C61.298 329.355 61.298 325.531 65.3832 323.172L233.548 226.082C235.455 224.982 238.545 224.982 240.452 226.082Z"
    id="float_server"
  ></path>
  <path
    fill="#5B6CA2"
    d="M408.896 332.123L241.489 428.775C240.013 429.68 238.557 430.033 236.934 430.033V464.518C238.904 464.518 239.366 464.169 241.489 463.233L408.896 366.58C411.372 365.292 412.125 363.262 412.312 361.317C412.312 361.317 412.312 326.583 412.312 327.722C412.312 328.86 411.42 330.514 408.896 332.123Z"
    id="float_server"
  ></path>
  <path
    fill="#6879AF"
    d="M240.92 429.077L255.155 420.857V432.434L251.511 439.064V457.432L241.489 463.242C240.116 463.858 239.141 464.518 236.934 464.518V430.024C238.695 430.024 239.862 429.701 240.92 429.077Z"
    id="float_server"
  ></path>
  <path
    fill="url(#paint4_linear_163_1030)"
    d="M65.084 331.984L232.379 428.571C233.882 429.619 235.101 430.005 236.934 430.005V464.523C234.656 464.523 234.285 464.215 232.379 463.214L65.084 366.442C62.4898 365 61.6417 362.992 61.6699 361.29V327.125C61.6899 329.24 62.4474 330.307 65.084 331.984Z"
    id="float_server"
  ></path>
  <path
    fill="#20273A"
    d="M400.199 361.032C403.195 359.302 405.623 355.096 405.623 351.637C405.623 348.177 403.195 346.775 400.199 348.505C397.203 350.235 394.775 354.441 394.775 357.9C394.775 361.359 397.203 362.762 400.199 361.032Z"
    id="float_server"
  ></path>
  <path
    fill="#20273A"
    d="M221.404 446.444C224.4 448.174 226.828 446.771 226.828 443.312C226.828 439.853 224.4 435.646 221.404 433.917C218.408 432.187 215.979 433.589 215.979 437.049C215.979 440.508 218.408 444.714 221.404 446.444Z"
    id="float_server"
  ></path>
  <path
    fill="#494F76"
    d="M102.895 359.589L97.9976 356.762V380.07L102.895 382.897V359.589Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M102.895 359.619L98.3394 356.989V379.854L102.895 382.484V359.619Z"
    id="float_server"
  ></path>
  <path
    fill="#494F76"
    d="M78.9793 345.923L74.0823 343.096V366.37L78.9793 369.198V345.923Z"
    id="float_server"
  ></path>
  <path
    fill="#494F76"
    d="M86.9512 350.478L82.0542 347.651V370.959L86.9512 373.787V350.478Z"
    id="float_server"
  ></path>
  <path
    fill="#494F76"
    d="M94.9229 355.034L90.0259 352.206V375.515L94.9229 378.342V355.034Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M86.951 350.509L82.3958 347.879V370.743L86.951 373.373V350.509Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M94.9227 355.064L90.3674 352.434V375.299L94.9227 377.929V355.064Z"
    class="estrobo_animation"
  ></path>
  <path
    fill="#313654"
    d="M78.9794 345.954L74.4241 343.324V366.188L78.9794 368.818V345.954Z"
    class="estrobo_animation"
  ></path>
  <path
    fill="#333B5F"
    d="M221.859 446.444C224.855 448.174 227.284 446.771 227.284 443.312C227.284 439.853 224.855 435.646 221.859 433.917C218.863 432.187 216.435 433.589 216.435 437.049C216.435 440.508 218.863 444.714 221.859 446.444Z"
    id="float_server"
  ></path>
  <path
    fill="#333B5F"
    d="M399.516 361.032C402.511 359.302 404.94 355.096 404.94 351.637C404.94 348.177 402.511 346.775 399.516 348.505C396.52 350.235 394.091 354.441 394.091 357.9C394.091 361.359 396.52 362.762 399.516 361.032Z"
    id="float_server"
  ></path>
  <path
    fill="#27C6FD"
    d="M88.3522 317.406L232.415 400.581C234.202 401.464 234.866 401.688 236.335 401.769V411.35C235.356 411.35 234.086 411.271 232.415 410.334L88.3522 327.159C86.3339 325.941 85.496 325.157 85.4707 323.256V313.486L88.3522 317.406Z"
    id="float_server"
  ></path>
  <path
    fill="#27C6FD"
    d="M384.318 317.504L240.254 400.679C238.914 401.444 238.295 401.688 236.335 401.769V411.35C237.315 411.35 238.704 411.27 240.236 410.333L384.318 327.257C386.457 326.15 387.151 325.303 387.258 323.287V313.976C386.768 315.446 386.21 316.354 384.318 317.504Z"
    id="float_server"
  ></path>
  <path
    stroke="url(#paint6_linear_163_1030)"
    fill="url(#paint5_linear_163_1030)"
    d="M240.452 169.141L408.617 266.231C412.703 268.59 412.703 272.414 408.617 274.772L240.452 371.862C238.545 372.962 235.455 372.962 233.548 371.862L65.3832 274.772C61.298 272.414 61.298 268.59 65.3832 266.231L233.548 169.141C235.455 168.04 238.545 168.04 240.452 169.141Z"
    id="float_server"
  ></path>
  <path
    fill="#5B6CA2"
    d="M408.896 275.182L241.489 371.834C240.013 372.739 238.557 373.092 236.934 373.092V407.577C238.904 407.577 239.366 407.229 241.489 406.292L408.896 309.64C411.372 308.352 412.125 306.321 412.312 304.376C412.312 304.376 412.312 269.642 412.312 270.781C412.312 271.92 411.42 273.573 408.896 275.182Z"
    id="float_server"
  ></path>
  <path
    fill="#6879AF"
    d="M240.92 372.135L255.155 363.915V375.493L251.511 382.123V400.491L241.489 406.3C240.116 406.916 239.141 407.577 236.934 407.577V373.083C238.695 373.083 239.862 372.759 240.92 372.135Z"
    id="float_server"
  ></path>
  <path
    fill="url(#paint7_linear_163_1030)"
    d="M65.084 275.043L232.379 371.63C233.882 372.678 235.101 373.064 236.934 373.064V407.582C234.656 407.582 234.285 407.274 232.379 406.273L65.084 309.501C62.4898 308.059 61.6417 306.051 61.6699 304.349V270.184C61.6899 272.299 62.4474 273.366 65.084 275.043Z"
    id="float_server"
  ></path>
  <path
    fill="#20273A"
    d="M400.199 304.091C403.195 302.362 405.623 298.155 405.623 294.696C405.623 291.237 403.195 289.835 400.199 291.564C397.203 293.294 394.775 297.5 394.775 300.959C394.775 304.419 397.203 305.821 400.199 304.091Z"
    id="float_server"
  ></path>
  <path
    fill="#20273A"
    d="M221.404 389.503C224.4 391.232 226.828 389.83 226.828 386.371C226.828 382.912 224.4 378.705 221.404 376.976C218.408 375.246 215.979 376.648 215.979 380.107C215.979 383.567 218.408 387.773 221.404 389.503Z"
    id="float_server"
  ></path>
  <path
    fill="#494F76"
    d="M102.553 301.281L97.656 298.454V321.762L102.553 324.59V301.281Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M102.553 301.312L97.9976 298.682V321.546L102.553 324.176V301.312Z"
    class="estrobo_animation"
  ></path>
  <path
    fill="#494F76"
    d="M78.6377 287.615L73.7407 284.788V308.063L78.6377 310.89V287.615Z"
    id="float_server"
  ></path>
  <path
    fill="#494F76"
    d="M86.6094 292.171L81.7124 289.343V312.652L86.6094 315.479V292.171Z"
    id="float_server"
  ></path>
  <path
    fill="#494F76"
    d="M94.5811 296.726L89.6841 293.899V317.207L94.5811 320.034V296.726Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M86.6095 292.201L82.0542 289.571V312.436L86.6095 315.066V292.201Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M94.5812 296.756L90.0259 294.126V316.991L94.5812 319.621V296.756Z"
    class="estrobo_animationV2"
  ></path>
  <path
    fill="#313654"
    d="M78.6376 287.646L74.0823 285.016V307.88L78.6376 310.51V287.646Z"
    id="float_server"
  ></path>
  <path
    fill="#333B5F"
    d="M221.859 389.503C224.855 391.232 227.284 389.83 227.284 386.371C227.284 382.912 224.855 378.705 221.859 376.976C218.863 375.246 216.435 376.648 216.435 380.107C216.435 383.567 218.863 387.773 221.859 389.503Z"
    id="float_server"
  ></path>
  <path
    fill="#333B5F"
    d="M399.516 304.091C402.511 302.362 404.94 298.155 404.94 294.696C404.94 291.237 402.511 289.835 399.516 291.564C396.52 293.294 394.091 297.5 394.091 300.959C394.091 304.419 396.52 305.821 399.516 304.091Z"
    id="float_server"
  ></path>
  <path
    fill="#27C6FD"
    d="M89.4907 214.912L233.554 298.087C235.341 298.97 236.003 299.194 237.474 299.275V308.856C236.494 308.856 235.223 308.777 233.554 307.84L89.4907 224.665C87.4726 223.447 86.6347 222.663 86.6094 220.762V210.993L89.4907 214.912Z"
    id="float_server"
  ></path>
  <path
    fill="#27C6FD"
    d="M385.457 215.01L241.393 298.185C240.053 298.951 239.434 299.194 237.474 299.275V308.856C238.454 308.856 239.844 308.776 241.375 307.839L385.457 224.763C387.597 223.656 388.29 222.809 388.397 220.793V211.482C387.907 212.953 387.349 213.86 385.457 215.01Z"
    id="float_server"
  ></path>
  <path
    fill="url(#paint8_linear_163_1030)"
    d="M66.1102 196.477L233.517 293.129C235.593 294.154 236.364 294.416 238.073 294.509V305.642C236.934 305.642 235.458 305.551 233.517 304.463L66.1102 207.81C63.7651 206.394 62.7914 205.483 62.762 203.275V191.922L66.1102 196.477Z"
    id="float_server"
  ></path>
  <path
    fill="#5B6CA2"
    d="M410.101 196.591L242.694 293.243C241.135 294.132 240.35 294.375 238.073 294.468V305.643C239.211 305.643 240.892 305.55 242.671 304.46L410.101 207.923C412.587 206.638 413.392 205.653 413.517 203.31V192.491C412.948 194.199 412.3 195.254 410.101 196.591Z"
    id="float_server"
  ></path>
  <path
    stroke="url(#paint10_linear_163_1030)"
    fill="url(#paint9_linear_163_1030)"
    d="M241.59 90.5623L409.756 187.652C413.842 190.011 413.842 193.835 409.756 196.194L241.59 293.284C239.684 294.384 236.593 294.384 234.687 293.284L66.5219 196.194C62.4367 193.835 62.4367 190.011 66.5219 187.652L234.687 90.5623C236.593 89.4616 239.684 89.4616 241.59 90.5623Z"
    id="float_server"
  ></path>
  <path
    fill="#20273A"
    d="M89.0427 195.334C92.0385 197.063 96.8956 197.063 99.8914 195.334C102.887 193.604 102.887 190.8 99.8914 189.07C96.8956 187.341 92.0385 187.341 89.0427 189.07C86.0469 190.8 86.0469 193.604 89.0427 195.334Z"
    id="float_server"
  ></path>
  <path
    fill="#20273A"
    d="M231.396 111.061C234.391 112.791 239.249 112.791 242.244 111.061C245.24 109.331 245.24 106.527 242.244 104.798C239.249 103.068 234.391 103.068 231.396 104.798C228.4 106.527 228.4 109.331 231.396 111.061Z"
    id="float_server"
  ></path>
  <path
    fill="#20273A"
    d="M374.887 194.195C377.883 195.925 382.74 195.925 385.736 194.195C388.732 192.465 388.732 189.661 385.736 187.932C382.74 186.202 377.883 186.202 374.887 187.932C371.891 189.661 371.891 192.465 374.887 194.195Z"
    id="float_server"
  ></path>
  <path
    fill="#20273A"
    d="M231.396 279.607C234.391 281.336 239.249 281.336 242.244 279.607C245.24 277.877 245.24 275.073 242.244 273.343C239.249 271.613 234.391 271.613 231.396 273.343C228.4 275.073 228.4 277.877 231.396 279.607Z"
    id="float_server"
  ></path>
  <path
    fill="#333B5F"
    d="M232.109 279.607C235.104 281.336 239.962 281.336 242.957 279.607C245.953 277.877 245.953 275.073 242.957 273.343C239.962 271.613 235.104 271.613 232.109 273.343C229.113 275.073 229.113 277.877 232.109 279.607Z"
    id="float_server"
  ></path>
  <path
    fill="#333B5F"
    d="M89.7563 195.334C92.7521 197.063 97.6092 197.063 100.605 195.334C103.601 193.604 103.601 190.8 100.605 189.07C97.6092 187.341 92.7521 187.341 89.7563 189.07C86.7605 190.8 86.7605 193.604 89.7563 195.334Z"
    id="float_server"
  ></path>
  <path
    fill="#333B5F"
    d="M232.109 111.061C235.104 112.791 239.962 112.791 242.957 111.061C245.953 109.331 245.953 106.527 242.957 104.798C239.962 103.068 235.104 103.068 232.109 104.798C229.113 106.527 229.113 109.331 232.109 111.061Z"
    id="float_server"
  ></path>
  <path
    fill="#333B5F"
    d="M375.6 194.195C378.595 195.925 383.453 195.925 386.448 194.195C389.444 192.465 389.444 189.661 386.448 187.932C383.453 186.202 378.595 186.202 375.6 187.932C372.604 189.661 372.604 192.465 375.6 194.195Z"
    id="float_server"
  ></path>
  <path
    stroke="#313654"
    d="M371.315 166.009L354.094 176.748C351.92 178.337 350.677 179.595 350.677 181.872L351.247 196.108C351.412 198.824 350.734 200.095 347.83 201.802L251.03 257.603C248.955 258.968 247.598 259.356 244.767 259.312L215.727 258.743C212.711 258.605 211.233 259.005 208.894 260.45L193.659 269.072"
    id="float_server"
  ></path>
  <path
    stroke="#313654"
    d="M345.691 151.204L328.328 161.374C326.154 162.963 324.911 164.221 324.911 166.498L325.481 180.734C325.646 183.45 324.968 184.721 322.064 186.428L225.264 242.229C223.19 243.594 221.832 243.982 219.001 243.938L189.961 243.369C186.946 243.231 185.468 243.631 183.128 245.076L167.124 253.698"
    id="float_server"
  ></path>
  <path
    stroke="#313654"
    d="M105.482 218.098L122.697 207.363C124.87 205.773 126.111 204.516 126.111 202.24L125.537 188.007C125.371 185.291 126.048 184.02 128.951 182.314L225.715 126.533C227.788 125.17 229.146 124.782 231.976 124.825L261.012 125.398C264.026 125.535 265.503 125.136 267.842 123.691L283.072 115.072"
    id="float_server"
  ></path>
  <path
    stroke="#313654"
    d="M131.121 232.893L148.482 222.725C150.656 221.136 151.898 219.879 151.898 217.601L151.327 203.367C151.162 200.65 151.839 199.379 154.743 197.673L251.531 141.878C253.605 140.514 254.962 140.126 257.794 140.17L286.832 140.74C289.847 140.878 291.325 140.478 293.664 139.032L309.667 130.412"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M327.961 242.79L301.907 227.748L300.673 228.46L326.727 243.503L327.961 242.79Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M354.625 227.426L328.56 212.377L327.326 213.09L353.392 228.139L354.625 227.426Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M300.864 258.519L274.707 243.417L273.474 244.129L299.631 259.231L300.864 258.519Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M176.498 155.101L150.21 139.924L148.977 140.636L175.264 155.813L176.498 155.101Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M193.703 145.191L167.388 129.998L166.154 130.711L192.469 145.903L193.703 145.191Z"
    id="float_server"
  ></path>
  <path
    fill="#313654"
    d="M158.333 165.69L131.974 150.472L130.74 151.184L157.099 166.402L158.333 165.69Z"
    id="float_server"
  ></path>
  <path
    fill="#20273A"
    d="M232.079 135.83C234.258 134.573 237.79 134.573 239.969 135.83L329.717 187.647C334.074 190.163 334.074 194.242 329.717 196.757L239.969 248.574C237.79 249.832 234.258 249.832 232.079 248.574L142.33 196.757C137.972 194.242 137.972 190.163 142.33 187.647L232.079 135.83Z"
    id="float_server"
  ></path>
  <path
    fill="url(#paint11_linear_163_1030)"
    d="M234.357 135.83C236.535 134.573 240.068 134.573 242.246 135.83L331.995 187.647C336.352 190.163 336.352 194.242 331.995 196.757L242.246 248.574C240.068 249.832 236.535 249.832 234.357 248.574L144.608 196.757C140.25 194.242 140.25 190.163 144.608 187.647L234.357 135.83Z"
    id="float_server"
  ></path>
  <path
    stroke-width="3"
    stroke="#27C6FD"
    d="M380.667 192.117V181.97C380.667 179.719 383.055 178.27 385.052 179.309L409.985 192.282C410.978 192.799 411.601 193.825 411.601 194.943V301.113C411.601 302.642 409.953 303.606 408.62 302.856L399.529 297.742"
    class="after_animation"
    id="float_server"
  ></path>
  <path
    stroke-width="3"
    stroke="#27C6FD"
    d="M94.7234 192.117V180.306C94.7234 179.214 94.1301 178.208 93.1744 177.68L70.5046 165.152C68.5052 164.047 66.0536 165.493 66.0536 167.778V185.326"
    id="float_server"
  ></path>
  <ellipse
    fill="#27C6FD"
    ry="1.50894"
    rx="1.50894"
    cy="192.117"
    cx="380.667"
    id="float_server"
  ></ellipse>
  <ellipse
    fill="#27C6FD"
    ry="1.50894"
    rx="1.50894"
    cy="192.117"
    cx="94.7235"
    id="float_server"
  ></ellipse>
  <ellipse
    fill="#27C6FD"
    ry="1.50894"
    rx="1.50894"
    cy="297.742"
    cx="399.529"
    id="float_server"
  ></ellipse>
  <ellipse
    fill="#27C6FD"
    ry="1.50894"
    rx="1.50894"
    cy="383.751"
    cx="221.474"
    id="float_server"
  ></ellipse>
  <ellipse
    fill="#27C6FD"
    ry="1.50894"
    rx="1.50894"
    cy="439.583"
    cx="221.474"
    id="float_server"
  ></ellipse>
  <path
    stroke-width="3"
    stroke="#27C6FD"
    d="M221.474 383.752L211.746 388.941C210.768 389.462 210.157 390.48 210.157 391.588V444.34C210.157 445.108 210.988 445.589 211.654 445.208L221.474 439.583"
    id="float_server"
  ></path>
  <path
    fill="url(#paint13_linear_163_1030)"
    d="M237.376 236.074L36 119.684V439.512C36.0957 441.966 36.7214 443.179 39.0056 445.021L200.082 538.547L231.362 556.441C233.801 557.806 235.868 558.222 237.376 558.328V236.074Z"
    id="layer_pared"
  ></path>
  <defs>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="556.454"
      x2="438.984"
      y1="235.918"
      x1="237.376"
      id="paint13_linear_163_1030"
    >
      <stop style="stop-color:#4457b3; stop-opacity:0" offset="10%"></stop>
      <stop style="stop-color:#4457b3; stop-opacity:1" offset="100%"></stop>
    </linearGradient>
  </defs>
  <path
    fill="url(#paint13_linear_163_1030)"
    d="M237.376 235.918L438.984 119.576V439.398C439.118 441.699 438.452 442.938 435.975 444.906L274.712 538.539L243.397 556.454C240.955 557.821 238.886 558.23 237.376 558.336V235.918Z"
    class="animatedStop"
    id="layer_pared"
  ></path>
  <defs>
    <filter
      color-interpolation-filters="sRGB"
      filterUnits="userSpaceOnUse"
      height="275.295"
      width="468.883"
      y="303.394"
      x="4.05835"
      id="filter0_i_163_1030"
    >
      <feFlood result="BackgroundImageFix" flood-opacity="0"></feFlood>
      <feBlend
        result="shape"
        in2="BackgroundImageFix"
        in="SourceGraphic"
        mode="normal"
      ></feBlend>
      <feColorMatrix
        result="hardAlpha"
        values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 127 0"
        type="matrix"
        in="SourceAlpha"
      ></feColorMatrix>
      <feOffset dy="4"></feOffset>
      <feGaussianBlur stdDeviation="2"></feGaussianBlur>
      <feComposite
        k3="1"
        k2="-1"
        operator="arithmetic"
        in2="hardAlpha"
      ></feComposite>
      <feColorMatrix
        values="0 0 0 0 0.220833 0 0 0 0 0.220833 0 0 0 0 0.220833 0 0 0 1 0"
        type="matrix"
      ></feColorMatrix>
      <feBlend
        result="effect1_innerShadow_163_1030"
        in2="shape"
        mode="normal"
      ></feBlend>
    </filter>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="336.055"
      x2="294.366"
      y1="60.1113"
      x1="135.05"
      id="paint0_linear_163_1030"
    >
      <stop stop-opacity="0.01" stop-color="white" offset="0.305"></stop>
      <stop stop-color="#4054B2" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="335.208"
      x2="180.935"
      y1="59.2405"
      x1="340.265"
      id="paint1_linear_163_1030"
    >
      <stop stop-opacity="0.01" stop-color="white" offset="0.305"></stop>
      <stop stop-color="#4054B2" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="420.619"
      x2="88.5367"
      y1="327.152"
      x1="412.313"
      id="paint2_linear_163_1030"
    >
      <stop stop-color="#313654"></stop>
      <stop stop-color="#313654" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="211.092"
      x2="168.239"
      y1="426.799"
      x1="236.934"
      id="paint3_linear_163_1030"
    >
      <stop stop-color="#7281B8"></stop>
      <stop stop-color="#333952" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="349.241"
      x2="232.379"
      y1="349.241"
      x1="65.0839"
      id="paint4_linear_163_1030"
    >
      <stop stop-color="#7281B8"></stop>
      <stop stop-color="#5D6EA4" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="363.678"
      x2="88.5367"
      y1="270.211"
      x1="412.313"
      id="paint5_linear_163_1030"
    >
      <stop stop-color="#313654"></stop>
      <stop stop-color="#313654" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="154.15"
      x2="168.239"
      y1="369.858"
      x1="236.934"
      id="paint6_linear_163_1030"
    >
      <stop stop-color="#7281B8"></stop>
      <stop stop-color="#333952" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="292.3"
      x2="232.379"
      y1="292.3"
      x1="65.0839"
      id="paint7_linear_163_1030"
    >
      <stop stop-color="#7281B8"></stop>
      <stop stop-color="#5D6EA4" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="198.899"
      x2="238.073"
      y1="198.899"
      x1="62.762"
      id="paint8_linear_163_1030"
    >
      <stop stop-color="#7382B9"></stop>
      <stop stop-color="#5D6EA4" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="191.599"
      x2="67.1602"
      y1="191.633"
      x1="413.451"
      id="paint9_linear_163_1030"
    >
      <stop stop-color="#5F6E99"></stop>
      <stop stop-color="#465282" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="191.599"
      x2="63.6601"
      y1="191.599"
      x1="417.16"
      id="paint10_linear_163_1030"
    >
      <stop stop-color="#7281B8"></stop>
      <stop stop-color="#333952" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="243.221"
      x2="156.734"
      y1="191.633"
      x1="335.442"
      id="paint11_linear_163_1030"
    >
      <stop stop-color="#313654"></stop>
      <stop stop-color="#313654" offset="1"></stop>
    </linearGradient>
    <linearGradient
      gradientUnits="userSpaceOnUse"
      y2="421.983"
      x2="-1.9283"
      y1="179.292"
      x1="138.189"
      id="paint12_linear_163_1030"
    >
      <stop stop-opacity="0.01" stop-color="white" offset="0.305"></stop>
      <stop stop-color="#4054B2" offset="1"></stop>
    </linearGradient>
  </defs>
</svg>

<style>
/* From Uiverse.io by https://uiverse.io/vajion_7943/selfish-hound-5 */
#svg_svg {
  zoom: 0.3;
}
.estrobo_animation {
  animation:
    floatAndBounce 4s infinite ease-in-out,
    strobe 0.8s infinite;
}

.estrobo_animationV2 {
  animation:
    floatAndBounce 4s infinite ease-in-out,
    strobev2 0.8s infinite;
}

#float_server {
  animation: floatAndBounce 4s infinite ease-in-out;
}

@keyframes floatAndBounce {
  0%,
  100% {
    transform: translateY(0);
  }

  50% {
    transform: translateY(-20px);
  }
}

@keyframes strobe {
  0%,
  50%,
  100% {
    fill: #17e300;
  }

  25%,
  75% {
    fill: #17e300b4;
  }
}

@keyframes strobev2 {
  0%,
  50%,
  100% {
    fill: rgb(255, 95, 74);
  }

  25%,
  75% {
    fill: rgb(255, 208, 1);
  }
}

/* Animación de los colores del gradiente */
@keyframes animateGradient {
  0% {
    stop-color: #19180f;
  }

  50% {
    stop-color: #ffdd00;
  }

  100% {
    stop-color: #e71a1a;
  }
}

/* Animación aplicada a los puntos del gradiente */
#paint13_linear_163_1030 stop {
  animation: animateGradient 4s infinite alternate;
}

</style
`;

  const loader3 = `
<!-- From Uiverse.io by https://uiverse.io/alexs_8179/splendid-squid-28 --> 
<svg xmlns="http://www.w3.org/2000/svg" height="200" width="200">
  <g style="order: -1; transform: scale(.75);">
    <polygon
      transform="rotate(45 100 100)"
      stroke-width="1"
      stroke="#17afbd"
      fill="none"
      points="70,70 148,50 130,130 50,150"
      id="bounce"
    ></polygon>
    <polygon
      transform="rotate(45 100 100)"
      stroke-width="1"
      stroke="#07e7fca4"
      fill="none"
      points="70,70 148,50 130,130 50,150"
      id="bounce2"
    ></polygon>
    <polygon
      transform="rotate(45 100 100)"
      stroke-width="2"
      stroke=""
      fill="#414750"
      points="70,70 150,50 130,130 50,150"
    ></polygon>
    <polygon
      stroke-width="2"
      stroke=""
      fill="url(#gradiente)"
      points="100,70 150,100 100,130 50,100"
    ></polygon>
    <defs>
      <linearGradient y2="100%" x2="10%" y1="0%" x1="0%" id="gradiente">
        <stop style="stop-color: #1e2026;stop-opacity:1" offset="20%"></stop>
        <stop style="stop-color:#414750;stop-opacity:1" offset="60%"></stop>
      </linearGradient>
    </defs>
    <polygon
      transform="translate(20, 31)"
      stroke-width="2"
      stroke=""
      fill="#227f8b"
      points="80,50 80,75 80,99 40,75"
    ></polygon>
    <polygon
      transform="translate(20, 31)"
      stroke-width="2"
      stroke=""
      fill="url(#gradiente2)"
      points="40,-40 80,-40 80,99 40,75"
    ></polygon>
    <defs>
      <linearGradient y2="100%" x2="0%" y1="-17%" x1="10%" id="gradiente2">
        <stop style="stop-color: #1f474400;stop-opacity:1" offset="20%"></stop>
        <stop
          style="stop-color:#10c6d354;stop-opacity:1"
          offset="100%"
          id="animatedStop"
        ></stop>
      </linearGradient>
    </defs>
    <polygon
      transform="rotate(180 100 100) translate(20, 20)"
      stroke-width="2"
      stroke=""
      fill="#17afbd"
      points="80,50 80,75 80,99 40,75"
    ></polygon>
    <polygon
      transform="rotate(0 100 100) translate(60, 20)"
      stroke-width="2"
      stroke=""
      fill="url(#gradiente3)"
      points="40,-40 80,-40 80,85 40,110.2"
    ></polygon>
    <defs>
      <linearGradient y2="100%" x2="10%" y1="0%" x1="0%" id="gradiente3">
        <stop style="stop-color: #10ccd300;stop-opacity:1" offset="20%"></stop>
        <stop
          style="stop-color:#d3a51054;stop-opacity:1"
          offset="100%"
          id="animatedStop"
        ></stop>
      </linearGradient>
    </defs>
    <polygon
      transform="rotate(45 100 100) translate(80, 95)"
      stroke-width="2"
      stroke=""
      fill="#ffffff"
      points="5,0 5,5 0,5 0,0"
      id="particles"
    ></polygon>
    <polygon
      transform="rotate(45 100 100) translate(80, 55)"
      stroke-width="2"
      stroke=""
      fill="#17afbd"
      points="6,0 6,6 0,6 0,0"
      id="particles"
    ></polygon>
    <polygon
      transform="rotate(45 100 100) translate(70, 80)"
      stroke-width="2"
      stroke=""
      fill="#17afbd"
      points="2,0 2,2 0,2 0,0"
      id="particles"
    ></polygon>
    <polygon
      stroke-width="2"
      stroke=""
      fill="#292d34"
      points="29.5,99.8 100,142 100,172 29.5,130"
    ></polygon>
    <polygon
      transform="translate(50, 92)"
      stroke-width="2"
      stroke=""
      fill="#1f2127"
      points="50,50 120.5,8 120.5,35 50,80"
    ></polygon>
  </g>
</svg>

<style>
/* From Uiverse.io by https://uiverse.io/alexs_8179/splendid-squid-28 */ 
@keyframes bounce {
  0%,
  100% {
    translate: 0px 36px;
  }
  50% {
    translate: 0px 46px;
  }
}
@keyframes bounce2 {
  0%,
  100% {
    translate: 0px 46px;
  }
  50% {
    translate: 0px 56px;
  }
}

@keyframes umbral {
  0% {
    stop-color: #10bcd32e;
  }
  50% {
    stop-color: #2fcad8;
  }
  100% {
    stop-color: #10d3982e;
  }
}
@keyframes partciles {
  0%,
  100% {
    translate: 0px 16px;
  }
  50% {
    translate: 0px 6px;
  }
}
#particles {
  animation: partciles 4s ease-in-out infinite;
}
#animatedStop {
  animation: umbral 4s infinite;
}
#bounce {
  animation: bounce 4s ease-in-out infinite;
  translate: 0px 36px;
}
#bounce2 {
  animation: bounce2 4s ease-in-out infinite;
  translate: 0px 46px;
  animation-delay: 0.5s;
}
</style>
`;

  const loader4 = `
<!-- From Uiverse.io by https://uiverse.io/KSAplay/silent-cobra-80 --> 
<div class="loader-ksaplay">
  <div class="cube">
    <div class="face middle front">
      <div class="cube cube-front">
        <div class="face front"></div>
        <div class="face back"></div>
        <div class="face left"></div>
        <div class="face right"></div>
        <div class="face top"></div>
        <div class="face bottom"></div>
      </div>
    </div>
    <div class="face middle back">
      <div class="cube cube-back">
        <div class="face front"></div>
        <div class="face back"></div>
        <div class="face left"></div>
        <div class="face right"></div>
        <div class="face top"></div>
        <div class="face bottom"></div>
      </div>
    </div>
    <div class="face middle left">
      <div class="cube cube-left">
        <div class="face front"></div>
        <div class="face back"></div>
        <div class="face left"></div>
        <div class="face right"></div>
        <div class="face top"></div>
        <div class="face bottom"></div>
      </div>
    </div>
    <div class="face middle right">
      <div class="cube cube-right">
        <div class="face front"></div>
        <div class="face back"></div>
        <div class="face left"></div>
        <div class="face right"></div>
        <div class="face top"></div>
        <div class="face bottom"></div>
      </div>
    </div>
    <div class="face middle top">
      <div class="cube cube-top">
        <div class="face front"></div>
        <div class="face back"></div>
        <div class="face left"></div>
        <div class="face right"></div>
        <div class="face top"></div>
        <div class="face bottom"></div>
      </div>
    </div>
    <div class="face middle bottom">
      <div class="cube cube-bottom">
        <div class="face front"></div>
        <div class="face back"></div>
        <div class="face left"></div>
        <div class="face right"></div>
        <div class="face top"></div>
        <div class="face bottom"></div>
      </div>
    </div>
  </div>
</div>

<style>
/* From Uiverse.io by https://uiverse.io/KSAplay/silent-cobra-80 */ 
.loader-ksaplay {
  position: relative;
  display: flex;
  align-items: center;
  justify-content: center;
  translate: 0 calc(var(--padding-screen) * -4);
}

.cube {
  position: absolute;
  width: 40px;
  transform-style: preserve-3d;
  transform: rotateX(-30deg) rotateY(45deg);
  transition: 300ms ease;
  cursor: pointer;
  animation: rotateCube 10s infinite linear;
}
/* change the distance between cubes with translateX */
.cube-front,
.cube-back {
  transform: translateX(40px) translateZ(-20px);
  animation: none;
}
/* change the distance between cubes with translateZ */
.cube-top,
.cube-bottom {
  transform: translateZ(20px);
  animation: none;
}
/* change the distance between cubes with translateX */
.cube-left,
.cube-right {
  transform: translateX(40px) translateZ(-20px);
  animation: none;
}

.face {
  position: absolute;
  transform-style: preserve-3d;
  width: 40px;
  height: 40px;
  background: rgb(21, 46, 75);
  background: radial-gradient(
    circle,
    rgba(21, 46, 75, 1) 0%,
    rgba(10, 14, 55, 1) 100%
  );
}

.front {
  transform: rotateY(0deg) translateZ(20px);
}

.back {
  transform: rotateY(180deg) translateZ(20px);
}

.left {
  transform: rotateY(-90deg) translateZ(20px);
}

.right {
  transform: rotateY(90deg) translateZ(20px);
}

.top {
  transform: rotateX(90deg) translateZ(20px);
}

.bottom {
  transform: rotateX(-90deg) translateZ(20px);
}

.cube-back:hover .face,
.cube-front:hover .face,
.cube-top:hover .face,
.cube-bottom:hover .face,
.cube-left:hover .face,
.cube-right:hover .face {
  background: rgb(255, 255, 255);
  background: radial-gradient(circle, white 60%, rgb(157, 208, 255) 100%);
  filter: drop-shadow(0px 0px 5px #e7faff)
    drop-shadow(0px 0px 15px rgb(75, 168, 255))
    drop-shadow(0px 0px 30px rgb(50, 156, 255));
}

.cube:active {
  transform: translateX(0px) translateZ(-20px);
}

.cube-back:active .face,
.cube-front:active .face,
.cube-top:active .face,
.cube-bottom:active .face,
.cube-left:active .face,
.cube-right:active .face {
  background: rgb(255, 255, 255);
  background: radial-gradient(circle, white 60%, rgb(157, 208, 255) 100%);
  filter: drop-shadow(0px 0px 5px #e7faff)
    drop-shadow(0px 0px 15px rgb(75, 168, 255))
    drop-shadow(0px 0px 30px rgb(50, 156, 255));
}

.middle {
  background: transparent;
}

@keyframes rotateCube {
  0% {
    transform: rotateX(-30deg) rotateY(45deg);
  }
  100% {
    transform: rotateX(-30deg) rotateY(405deg);
  }
}
</style>
`;

  const loader5 = `
<!-- From Uiverse.io by https://uiverse.io/Z4drus/average-lizard-53 --> 
<div class="loader-Z4drus">
  <div class="crystal"></div>
  <div class="crystal"></div>
  <div class="crystal"></div>
  <div class="crystal"></div>
  <div class="crystal"></div>
  <div class="crystal"></div>
</div>
<style>
/* From Uiverse.io by Z4drus */ 
.loader-z4drus {
  position: relative;
  width: 200px;
  height: 200px;
  perspective: 800px;
  transform: translate(50px, -50px) !important;
}

.crystal {
  position: absolute;
  top: 50%;
  left: 50%;
  width: 60px;
  height: 60px;
  opacity: 0;
  transform-origin: bottom center;
  transform: translate(-50%, -50%) rotateX(45deg) rotateZ(0deg);
  animation: spin 4s linear infinite, emerge 2s ease-in-out infinite alternate,
    fadeIn 0.3s ease-out forwards;
  border-radius: 10px;
  visibility: hidden;
}

@keyframes spin {
  from {
    transform: translate(-50%, -50%) rotateX(45deg) rotateZ(0deg);
  }
  to {
    transform: translate(-50%, -50%) rotateX(45deg) rotateZ(360deg);
  }
}

@keyframes emerge {
  0%,
  100% {
    transform: translate(-50%, -50%) scale(0.5);
    opacity: 0;
  }
  50% {
    transform: translate(-50%, -50%) scale(1);
    opacity: 1;
  }
}

@keyframes fadeIn {
  to {
    visibility: visible;
    opacity: 0.8;
  }
}

.crystal:nth-child(1) {
  background: linear-gradient(45deg, #003366, #336699);
  animation-delay: 0s;
}

.crystal:nth-child(2) {
  background: linear-gradient(45deg, #003399, #3366cc);
  animation-delay: 0.3s;
}

.crystal:nth-child(3) {
  background: linear-gradient(45deg, #0066cc, #3399ff);
  animation-delay: 0.6s;
}

.crystal:nth-child(4) {
  background: linear-gradient(45deg, #0099ff, #66ccff);
  animation-delay: 0.9s;
}

.crystal:nth-child(5) {
  background: linear-gradient(45deg, #33ccff, #99ccff);
  animation-delay: 1.2s;
}

.crystal:nth-child(6) {
  background: linear-gradient(45deg, #66ffff, #ccffff);
  animation-delay: 1.5s;
}
</style>
`;

  const loader6 = `
<!-- From Uiverse.io by https://uiverse.io/mobinkakei/dry-insect-94 --> 
<span class="loader-mobinkakei"></span>

<style>
.loader-mobinkakei {
  position: absolute;
  bottom: 0px;
  left: 0px;
  width: 60px;
  height: 60px;
  background: #a19dad;
  transform: rotateX(60deg) rotate(45deg);
  color: #fff;
  animation: layers1mobinkakei 1s linear infinite alternate;
}
.loader-mobinkakei:after {
  content: "";
  position: absolute;
  inset: 0;
  background: rgba(255, 255, 255, 0.7);
  animation: layerTrmobinkakei 1s linear infinite alternate;
}

@keyframes layers1mobinkakei {
  0% {
    box-shadow: 0px 0px 0 0px;
  }
  90%,
  100% {
    box-shadow: 20px 20px 0 -4px;
  }
}
@keyframes layerTrmobinkakei {
  0% {
    transform: translate(0, 0) scale(1);
  }
  100% {
    transform: translate(-25px, -25px) scale(1);
  }
}
</style>
`;

  const loader7 = `
  <!-- https://codepen.io/alexanderkwright/pen/xbKqPQ -->
<div class='container-alexanderkwright'>
  <i class='layer-alexanderkwright'></i>
  <i class='layer-alexanderkwright'></i>
  <i class='layer-alexanderkwright'></i>
</div>
<style>
/**
 * Create the loop delay with
 * the extra keyframes
 */
@-webkit-keyframes moveup {
  0%, 60%, 100% {
    transform: rotateX(50deg) rotateY(0deg) rotateZ(45deg) translateZ(0);
  }
  25% {
    transform: rotateX(50deg) rotateY(0deg) rotateZ(45deg) translateZ(1em);
  }
}
@keyframes moveup {
  0%, 60%, 100% {
    transform: rotateX(50deg) rotateY(0deg) rotateZ(45deg) translateZ(0);
  }
  25% {
    transform: rotateX(50deg) rotateY(0deg) rotateZ(45deg) translateZ(1em);
  }
}
@-webkit-keyframes movedown {
  0%, 60%, 100% {
    transform: rotateX(50deg) rotateY(0deg) rotateZ(45deg) translateZ(0);
  }
  25% {
    transform: rotateX(50deg) rotateY(0deg) rotateZ(45deg) translateZ(-1em);
  }
}
@keyframes movedown {
  0%, 60%, 100% {
    transform: rotateX(50deg) rotateY(0deg) rotateZ(45deg) translateZ(0);
  }
  25% {
    transform: rotateX(50deg) rotateY(0deg) rotateZ(45deg) translateZ(-1em);
  }
}
/**
 * Square layer styles
 */
.layer-alexanderkwright {
  display: block;
  position: absolute;
  height: 3em;
  width: 3em;
  box-shadow: 3px 3px 2px rgba(0, 0, 0, 0.2);
  transform: rotateX(50deg) rotateY(0deg) rotateZ(45deg);
}
.layer-alexanderkwright:nth-of-type(1) {
  background: #534a47;
  margin-top: 1.5em;
  -webkit-animation: movedown 1.8s cubic-bezier(0.39, 0.575, 0.565, 1) 0.9s infinite normal;
          animation: movedown 1.8s cubic-bezier(0.39, 0.575, 0.565, 1) 0.9s infinite normal;
}
.layer-alexanderkwright:nth-of-type(1):before {
  content: "";
  position: absolute;
  width: 85%;
  height: 85%;
  background: #37332f;
}
.layer-alexanderkwright:nth-of-type(2) {
  background: #5a96bc;
  margin-top: 0.75em;
}
.layer-alexanderkwright:nth-of-type(3) {
  background: rgba(255, 255, 255, 0.6);
  -webkit-animation: moveup 1.8s cubic-bezier(0.39, 0.575, 0.565, 1) infinite normal;
          animation: moveup 1.8s cubic-bezier(0.39, 0.575, 0.565, 1) infinite normal;
}

/* Stage and link styles */

.container-alexanderkwright {
  position: absolute;
  bottom: calc(var(--padding-screen) * 2);
  left: 0;
}

.link {
  position: absolute;
  top: 30%;
  left: 50%;
  color: rgba(255, 255, 255, 0.5);
  font: 400 1em Helvetica Neue, Helvetica, sans-serif;
  transform: translate(-50%, -50%);
}
.link a {
  color: #ea4c89;
  text-decoration: none;
}
  </style>
`

  const loader8 = `
<!-- From Uiverse.io by https://uiverse.io/Juanes200122/fluffy-lizard-32 --> 
<div class="container_SevMini">
  <div class="SevMini">
    <svg
      width="74"
      height="90"
      viewBox="0 0 74 90"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      <path
        d="M40 76.5L72 57V69.8615C72 70.5673 71.628 71.2209 71.0211 71.5812L40 90V76.5Z"
        fill="#396CAA"
      ></path>
      <path
        d="M34 75.7077L2 57V69.8615C2 70.5673 2.37203 71.2209 2.97892 71.5812L34 90V75.7077Z"
        fill="#396DAC"
      ></path>
      <path d="M34 76.5H40V90H34V76.5Z" fill="#396CAA"></path>
      <path
        d="M3.27905 55.593L35.2806 37.5438C36.3478 36.9419 37.6522 36.9419 38.7194 37.5438L70.721 55.593C71.7294 56.1618 71.7406 57.6102 70.7411 58.1945L39.2712 76.593C37.8682 77.4133 36.1318 77.4133 34.7288 76.593L3.25887 58.1945C2.25937 57.6102 2.27061 56.1618 3.27905 55.593Z"
        fill="#163C79"
        stroke="#396CAA"
      ></path>
      <path
        d="M40 79L72 60V70.4001C72 71.1151 71.6183 71.7758 70.9987 72.1329L40 90V79Z"
        fill="#173D7A"
      ></path>
      <path d="M34 79L3 61V71.5751L34 90V79Z" fill="#0665B2"></path>
      <path
        id="strobe_color1"
        d="M58 72.5L60.5 71V74L58 75.5V72.5Z"
        fill="#FF715E"
      ></path>
      <path
        id="strobe_color2"
        d="M63 69.5L65.5 68V71L63 72.5V69.5Z"
        fill="#17e300b4"
      ></path>
      <path d="M68 66.5L70.5 65V68L68 69.5V66.5Z" fill="#FF715E"></path>
      <path
        d="M40 58.5L72 39V51.8615C72 52.5673 71.628 53.2209 71.0211 53.5812L40 72V58.5Z"
        fill="#396CAA"
      ></path>
      <path
        d="M34 57.7077L2 39V51.8615C2 52.5673 2.37203 53.2209 2.97892 53.5812L34 72V57.7077Z"
        fill="#396DAC"
      ></path>
      <path d="M34 58.5H40V72H34V58.5Z" fill="#396CAA"></path>
      <path
        d="M3.27905 37.593L35.2806 19.5438C36.3478 18.9419 37.6522 18.9419 38.7194 19.5438L70.721 37.593C71.7294 38.1618 71.7406 39.6102 70.7411 40.1945L39.2712 58.593C37.8682 59.4133 36.1318 59.4133 34.7288 58.593L3.25887 40.1945C2.25937 39.6102 2.27061 38.1618 3.27905 37.593Z"
        fill="#163C79"
        stroke="#396CAA"
      ></path>
      <path
        d="M40 61L72 42V52.4001C72 53.1151 71.6183 53.7758 70.9987 54.1329L40 72V61Z"
        fill="#173D7A"
      ></path>
      <path d="M34 61L3 43V53.5751L34 72V61Z" fill="#0665B2"></path>
      <path d="M58 54.5L60.5 53V56L58 57.5V54.5Z" fill="#FF715E"></path>
      <path d="M63 51.5L65.5 50V53L63 54.5V51.5Z" fill="black"></path>
      <path
        id="strobe_color1"
        d="M63 51.5L65.5 50V53L63 54.5V51.5Z"
        fill="#FF715E"
      ></path>
      <path d="M68 48.5L70.5 47V50L68 51.5V48.5Z" fill="#FF715E"></path>
      <path
        d="M40 40.5L72 21V33.8615C72 34.5673 71.628 35.2209 71.0211 35.5812L40 54V40.5Z"
        fill="#396CAA"
      ></path>
      <path
        d="M34 39.7077L2 21V33.8615C2 34.5673 2.37203 35.2209 2.97892 35.5812L34 54V39.7077Z"
        fill="#396DAC"
      ></path>
      <path d="M34 40.5H40V54H34V40.5Z" fill="#396CAA"></path>
      <path
        d="M3.27905 19.593L35.2806 1.54381C36.3478 0.941872 37.6522 0.941872 38.7194 1.54381L70.721 19.593C71.7294 20.1618 71.7406 21.6102 70.7411 22.1945L39.2712 40.593C37.8682 41.4133 36.1318 41.4133 34.7288 40.593L3.25887 22.1945C2.25937 21.6102 2.27061 20.1618 3.27905 19.593Z"
        fill="#124E89"
        stroke="#396CAA"
      ></path>
      <path
        d="M40 43L72 24V34.4001C72 35.1151 71.6183 35.7758 70.9987 36.1329L40 54V43Z"
        fill="#173D7A"
      ></path>
      <path d="M34 43L3 25V35.5751L34 54V43Z" fill="#0665B2"></path>
      <path d="M68 30.5L70.5 29V32L68 33.5V30.5Z" fill="#FF715E"></path>
      <path
        id="strobe_color3"
        d="M58 36.5L60.5 35V38L58 39.5V36.5Z"
        fill="#FF715E"
      ></path>
      <path d="M63 33.5L65.5 32V35L63 36.5V33.5Z" fill="#FF715E"></path>
      <path
        d="M20.1902 22.0719C18.8101 21.3026 18.8252 19.3119 20.2168 18.5636L36.1054 10.0189C37.2884 9.3827 38.7116 9.3827 39.8946 10.0189L55.7832 18.5636C57.1748 19.3119 57.1899 21.3026 55.8098 22.0719L40.4345 30.6429C38.9211 31.4865 37.0789 31.4865 35.5655 30.6429L20.1902 22.0719Z"
        fill="#396CAA"
      ></path>
      <path
        d="M11 52.755C11 51.9801 11.8432 51.4997 12.5098 51.8947L23.5196 58.419C24.1273 58.7792 24.5 59.4332 24.5 60.1396V60.245C24.5 61.0199 23.6568 61.5003 22.9902 61.1053L11.9804 54.581C11.3727 54.2208 11 53.5668 11 52.8604V52.755Z"
        fill="#396CAA"
      ></path>
      <mask
        id="mask0_2_176"
        style="mask-type:alpha"
        maskUnits="userSpaceOnUse"
        x="11"
        y="51"
        width="14"
        height="11"
      >
        <path
          d="M11 52.755C11 51.9801 11.8432 51.4997 12.5098 51.8947L23.5196 58.419C24.1273 58.7792 24.5 59.4332 24.5 60.1396V60.245C24.5 61.0199 23.6568 61.5003 22.9902 61.1053L11.9804 54.581C11.3727 54.2208 11 53.5668 11 52.8604V52.755Z"
          fill="#396CAA"
        ></path>
      </mask>
      <g mask="url(#mask0_2_176)">
        <path
          d="M11.5 52.7417C11.5 51.9803 12.3349 51.5138 12.9833 51.9128L23.5482 58.4143C24.1397 58.7783 24.5 59.4231 24.5 60.1176V61.5L12.4598 54.4195C11.8651 54.0698 11.5 53.4315 11.5 52.7417V52.7417Z"
          fill="#163874"
        ></path>
      </g>
      <mask
        id="mask1_2_176"
        style="mask-type:alpha"
        maskUnits="userSpaceOnUse"
        x="19"
        y="9"
        width="38"
        height="23"
      >
        <path
          d="M20.1902 22.0719C18.8101 21.3026 18.8252 19.3119 20.2168 18.5636L36.1054 10.0189C37.2884 9.3827 38.7116 9.3827 39.8946 10.0189L55.7832 18.5636C57.1748 19.3119 57.1899 21.3026 55.8098 22.0719L40.4345 30.6429C38.9211 31.4865 37.0789 31.4865 35.5655 30.6429L20.1902 22.0719Z"
          fill="#396CAA"
        ></path>
      </mask>
      <g mask="url(#mask1_2_176)">
        <path
          d="M18 21.3115L36.167 11.9451C37.3171 11.3521 38.6829 11.3521 39.833 11.9451L58 21.3115L40.3567 30.7405C38.8841 31.5275 37.1159 31.5275 35.6433 30.7405L18 21.3115Z"
          fill="#173D7A"
        ></path>
      </g>
      <path
        d="M37.447 21.565L35 19.9799L37.6941 18.66L40.141 20.245L37.447 21.565Z"
        fill="#FF715E"
      ></path>
      <path
        d="M48.9738 30.8646L47.0741 29.7745L49.1792 28.684L51.0789 29.7741L48.9738 30.8646Z"
        fill="#173E7B"
      ></path>
      <path
        d="M52.0661 29.0093L50.1635 27.9242L52.2657 26.8282L54.1682 27.9133L52.0661 29.0093Z"
        fill="#173E7B"
      ></path>
      <path
        id="strobe_led1"
        d="M55.1521 27.1464L53.2538 26.054L55.3602 24.9661L57.2585 26.0586L55.1521 27.1464Z"
        fill="#3A6DAB"
      ></path>
    </svg>
  </div>
  <div class="Ghost">
    <svg
      width="60"
      height="36"
      viewBox="0 0 60 36"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      <path
        d="M1.96545 19.4296C0.643777 18.6484 0.658726 16.7309 1.99242 15.9705L28.0186 1.12982C29.2467 0.429534 30.7533 0.429533 31.9814 1.12982L58.0076 15.9704C59.3413 16.7309 59.3562 18.6484 58.0346 19.4296L32.5442 34.4962C30.9749 35.4238 29.0251 35.4238 27.4558 34.4962L1.96545 19.4296Z"
        fill="#3C4F6D"
      ></path>
    </svg>
  </div>
</div>
<style>
/* From Uiverse.io by Juanes200122 */ 
.container_SevMini {
  display: flex;
  justify-content: center;
  align-items: center;
  flex-direction: column;
}

.Ghost {
  transform: translate(0px, -25px);
  z-index: -1;
  animation: opacidad 4s infinite ease-in-out;
}

@keyframes opacidad {
  0% {
    opacity: 1;
    scale: 1;
  }

  50% {
    opacity: 0.5;
    scale: 0.9;
  }

  100% {
    opacity: 1;
    scale: 1;
  }
}

@keyframes estroboscopico {
  0% {
    opacity: 1;
  }

  50% {
    opacity: 0;
  }

  51% {
    opacity: 1;
  }

  100% {
    opacity: 1;
  }
}

@keyframes rebote {
  0%,
  100% {
    transform: translateY(0);
  }

  50% {
    transform: translateY(-10px);
  }
}

@keyframes estroboscopico1 {
  0%,
  50%,
  100% {
    fill: rgb(255, 95, 74);
  }

  25%,
  75% {
    fill: rgb(16, 53, 115);
  }
}

@keyframes estroboscopico2 {
  0%,
  50%,
  100% {
    fill: #17e300;
  }

  25%,
  75% {
    fill: #17e300b4;
  }
}

.SevMini {
  animation: rebote 4s infinite ease-in-out;
}

#strobe_led1 {
  animation: estroboscopico 0.5s infinite;
}

#strobe_color1 {
  animation: estroboscopico2 0.8s infinite;
}

#strobe_color3 {
  animation: estroboscopico1 0.8s infinite;
  animation-delay: 3s;
}
</style>
`

  const loader9 = `
<!-- https://codepen.io/florian-gropp/pen/BOMyae -->
<svg id="svg-sprite">
	<symbol id="paw" viewBox="0 0 249 209.32">
		<ellipse cx="27.917" cy="106.333" stroke-width="0" rx="27.917" ry="35.833"/>
		<ellipse cx="84.75" cy="47.749" stroke-width="0" rx="34.75" ry="47.751"/>
		<ellipse cx="162" cy="47.749" stroke-width="0" rx="34.75" ry="47.751"/>
		<ellipse cx="221.083" cy="106.333" stroke-width="0" rx="27.917" ry="35.833"/>
		<path stroke-width="0" d="M43.98 165.39s9.76-63.072 76.838-64.574c0 0 71.082-6.758 83.096 70.33 0 0 2.586 19.855-12.54 31.855 0 0-15.75 17.75-43.75-6.25 0 0-7.124-8.374-24.624-7.874 0 0-12.75-.125-21.5 6.625 0 0-16.375 18.376-37.75 12.75 0 0-28.29-7.72-19.77-42.86z"/>
	</symbol>
</svg>

<div class="ajax-loader">
	<div class="paw"><svg class="icon"><use xlink:href="#paw" /></svg></div>
	<div class="paw"><svg class="icon"><use xlink:href="#paw" /></svg></div>
	<div class="paw"><svg class="icon"><use xlink:href="#paw" /></svg></div>
	<div class="paw"><svg class="icon"><use xlink:href="#paw" /></svg></div>
	<div class="paw"><svg class="icon"><use xlink:href="#paw" /></svg></div>
	<div class="paw"><svg class="icon"><use xlink:href="#paw" /></svg></div>
</div>
<style>
.ajax-loader {
  position: absolute;
  bottom: 50px;
  left: 200px;
  transform-origin: 50% 50%;
  transform: rotate(90deg) translate(-50%, 0%);
  font-size: 50px;
  width: 1em;
  height: 3em;
  color: #d31145;
}
.ajax-loader .paw {
  width: 1em;
  height: 1em;
  -webkit-animation: 2050ms pawAnimation ease-in-out infinite;
          animation: 2050ms pawAnimation ease-in-out infinite;
  opacity: 0;
}
.ajax-loader .paw svg {
  width: 100%;
  height: 100%;
}
.ajax-loader .paw .icon {
  fill: currentColor;
}
.ajax-loader .paw:nth-child(odd) {
  transform: rotate(-10deg);
}
.ajax-loader .paw:nth-child(even) {
  transform: rotate(10deg) translate(125%, 0);
}
.ajax-loader .paw:nth-child(1) {
  -webkit-animation-delay: 1.25s;
          animation-delay: 1.25s;
}
.ajax-loader .paw:nth-child(2) {
  -webkit-animation-delay: 1s;
          animation-delay: 1s;
}
.ajax-loader .paw:nth-child(3) {
  -webkit-animation-delay: 0.75s;
          animation-delay: 0.75s;
}
.ajax-loader .paw:nth-child(4) {
  -webkit-animation-delay: 0.5s;
          animation-delay: 0.5s;
}
.ajax-loader .paw:nth-child(5) {
  -webkit-animation-delay: 0.25s;
          animation-delay: 0.25s;
}
.ajax-loader .paw:nth-child(6) {
  -webkit-animation-delay: 0s;
          animation-delay: 0s;
}
.no-cssanimations .ajax-loader .paw {
  opacity: 1;
}

@-webkit-keyframes pawAnimation {
  0% {
    opacity: 1;
  }
  50% {
    opacity: 0;
  }
  100% {
    opacity: 0;
  }
}

@keyframes pawAnimation {
  0% {
    opacity: 1;
  }
  50% {
    opacity: 0;
  }
  100% {
    opacity: 0;
  }
}
</style>
`;

  const loader10 = `
<!-- https://codepen.io/stivaliserna/pen/jObPyKe -->
<svg class="sausage-dog-animation" xmlns="http://www.w3.org/2000/svg" viewBox="-50 0 1200 1080">
  <ellipse class="shadow" ry="45" rx="350" cy="816" cx="498" opacity="1" fill="#B2CAE8" fill-opacity="1" stroke="#B2CAE8" stroke-width="4" />
  <path class="tail-blur" fill="#6f5a4b" d="M 180.265,568.972 14.092,504.432 C 40.893351,428.54412 92.941075,394.6756 159.419,390.74 l 29.972,170.684 c 1.155,6.575 -2.931,9.954 -9.126,7.548 z" opacity=".296" />
  <path class="tail" fill="#6f5a4b" stroke="#6f5a4b" stroke-width="12" d="m 161.6285,568.63016 20.92664,-20.00034 C 151.50961,521.73829 14.092,504.432 14.092,504.432 c 0,0 128.8135,26.71916 147.5365,64.19816 z" />
  <g class="back-legs">
    <path fill="#9b6e34" d="M180.059 589.035h121.038l-23.651 219.517-58.432 2.217-38.955-221.734z" />
    <path fill="#9b6e34" stroke="#9b6e34" stroke-width="6" d="M270.996 760.244c-28.22 0-51.088 22.555-51.536 50.525h103.071c-.447-27.97-23.315-50.525-51.535-50.525z" />
    <path fill="#bd8b4a" d="M206.036 589.035h121.039l-23.651 219.517-58.433 2.217-38.955-221.734z" />
    <path fill="#bd8b4a" stroke="#bd8b4a" stroke-width="6" d="M296.973 760.244c-28.219 0-51.088 22.555-51.535 50.525h103.071c-.447-27.97-23.316-50.525-51.536-50.525z" />
    <path fill="#9b6e34" fill-opacity="1" stroke="#9b6e34" stroke-width="1" d="m 341.34945,810.26154 c -1.48174,-30.60183 -20.2921,-36.5324 -23.61393,-35.8329 -1.37117,0.0393 21.53672,10.53459 17.96335,35.34867 z" />
  </g>
  <g class="front-legs">
    <path class="leg" fill="#9b6e34" d="m 640.90169,580.10511 119.14374,21.32995 -61.96529,211.91363 -57.90822,-8.11489 z" />
    <path fill="#9b6e34" stroke="#9b6e34" stroke-width="6" d="m 711.354,758.244 c -28.22,0 -51.089,22.555 -51.536,50.525 h 103.071 c -0.447,-27.97 -23.316,-50.525 -51.535,-50.525 z" />
    <path fill="#bd8b4a" fill-opacity="1" stroke="#bd8b4a" stroke-width="1" d="m 760.35315,810.35988 c -1.48174,-30.60183 -20.2921,-36.5324 -23.61393,-35.8329 -1.37117,0.0393 21.53672,10.53459 17.96335,35.34867 z" />
    <path class="leg" fill="#bd8b4a" d="m 619.46642,579.5532 118.83259,23.00558 -64.94292,211.02015 -57.78921,-8.92963 z" />
    <path fill="#bd8b4a" stroke="#bd8b4a" stroke-width="6" d="m 688.228,758.24364 c -56.81455,-1.81204 -59.84071,0.90112 -51.535,50.525 h 103.071 c -0.447,-27.97 -23.33034,-49.62541 -51.536,-50.525 z" />
    <path fill="#9b6e34" fill-opacity="1" stroke="#9b6e34" stroke-width="1" d="m 734.97661,810.43045 c -1.48174,-30.60183 -20.2921,-36.5324 -23.61393,-35.8329 -1.37117,0.0393 21.53672,10.53459 17.96335,35.34867 z" />
  </g>
  <g class="lean">
    <g class="body">
      <path fill="#6f5a4b" stroke="#6f5a4b" stroke-width="3.25" d="M751.059 361.906c0-29.802 45.436-69.614 45.436-69.614s49.234-43.162 87.533-43.162c73.437 0 111.791 39.402 111.791 101.686 0 31.044-9.528 61.914-28.18 85.07-18.769 23.303-46.777 38.795-83.611 38.795-73.437 0-132.969-50.491-132.969-112.775z" />
      <path fill="#6f5a4b" stroke="#6f5a4b" stroke-width="3.25" d="M161.724 552.442c43.998-51.241 266.636-59.779 370.667-66.522 85.958-5.571 279.053-236.247 279.053-163.339 0 53.584 51.308 278.307-16.195 344.713-68.328 67.217-165.195 62.163-239.502 61.231-52.92-.663-131.793-11.253-226.832-15.838-77.24-3.727-120.165-11.768-152.525-38.198-41.065-33.54-39.869-92.695-14.666-122.047z" />
      <path fill="#695445" stroke="#695445" stroke-width="2.06" d="M412.368 528.039l22.614-14.883c5.4-3.554 14.781-6.181 20.954-5.866l13.681.698c6.173.314 15.487 2.785 20.804 5.518l20.246 10.406c5.317 2.733 11.558 9.379 13.94 14.844l2.905 6.668c2.382 5.465 2.362 14.665-.044 20.548l-1.537 3.759c-2.406 5.883-8.735 13.532-14.136 17.084l-10.784 7.093c-5.401 3.552-13.927 9.649-19.044 13.617l-11.314 8.774c-5.116 3.969-14.146 8.913-20.168 11.043l-2.382.842c-6.022 2.13-15.885 3.458-22.029 2.966l-7.054-.565c-6.145-.492-16.135-1.112-22.313-1.384l-6.765-.298c-6.178-.273-15.545-2.626-20.922-5.256l-.197-.096c-5.376-2.63-9.598-9.753-9.43-15.91l.316-11.56c.169-6.157 3.685-15.154 7.854-20.095l27.479-32.565c4.169-4.941 11.926-11.828 17.326-15.382z" />
      <path fill="#695445" stroke="#695445" stroke-width="5" d="M392.18 675.438l8.165-10.299c1.928-2.433 5.608-5.685 8.218-7.265l15.541-9.405c2.61-1.58 7.112-2.313 10.055-1.636l13.58 3.123c2.943.676 7.135 2.952 9.364 5.081l.26.249c2.229 2.13 6.286 2.86 9.063 1.631l4.337-1.919c2.776-1.229 7.431-2.671 10.398-3.221l2.133-.396c2.966-.55 7.647-.066 10.455 1.082l3.466 1.417c2.807 1.147 7.4 2.889 10.258 3.891l1.012.354c2.858 1.002 7.401 2.867 10.148 4.166l5.2 2.46c2.747 1.299 5.746 4.788 6.699 7.792l3.367 10.618c.953 3.003.26 7.492-1.548 10.025l-2.543 3.565c-1.808 2.533-5.672 5.067-8.631 5.66l-5.189 1.039c-2.96.592-7.792.868-10.795.615l-8.061-.678c-3.003-.252-7.719-1.371-10.534-2.498l-5.712-2.287c-2.815-1.127-7.323-.988-10.07.312l-5.201 2.46c-2.747 1.299-7.391 2.715-10.373 3.163l-5.106.767c-2.983.448-7.717-.001-10.574-1.002l-1.013-.355c-2.857-1.001-6.663-3.848-8.499-6.357l-.923-1.262c-1.837-2.509-5.753-4.827-8.747-5.177l-2.791-.326c-2.994-.35-7.652.409-10.405 1.695l-3.665 1.713c-2.753 1.286-6.27.145-7.855-2.547l-4.104-6.969c-1.586-2.692-1.308-6.847.62-9.279z" />
      <path fill="#695445" stroke="#695445" stroke-width="5" d="M559.503 565.641l-6.595 16.23c-1.842 4.533-3.19 12.298-3.012 17.344l.375 10.576c.179 5.045 1.874 13.072 3.787 17.928l3.273 8.314c1.913 4.856 4.607 12.825 6.018 17.799l2.271 8.005c1.411 4.973 5.602 12.104 9.361 15.927l.137.14c3.759 3.823 10.993 6.882 16.159 6.833l1.542-.015c5.166-.049 11.579-3.345 14.325-7.361l1.801-2.633c2.746-4.016 4.7-11.368 4.364-16.42l-.529-7.952c-.336-5.053-1.859-13.16-3.402-18.108l-2.384-7.645c-1.543-4.948-2.962-13.051-3.17-18.098l-.145-3.513c-.208-5.047 2.286-12.052 5.569-15.646l6.804-7.445c3.283-3.594 7.962-9.9 10.45-14.085l1.031-1.735c2.488-4.185 3.56-11.641 2.394-16.653l-.216-.927c-1.166-5.012-6.322-9.396-11.516-9.793l-10.284-.786c-5.194-.397-12.79 1.392-16.967 3.995l-20.542 12.802c-4.177 2.604-9.056 8.389-10.899 12.922z" />
      <path fill="#695445" stroke="#695445" stroke-width="5" d="M642.333 590.4l-3.736 1.374c-3.152 1.16-7.235 4.448-9.12 7.344l-3.019 4.638c-1.885 2.896-3.749 8.067-4.163 11.55l-.015.127c-.415 3.483 1.926 6.728 5.228 7.248l3.191.503c3.302.521 7.946-1.01 10.374-3.42l4.843-4.808c2.427-2.41 5.87-6.749 7.689-9.691l3.257-5.267c1.82-2.943.874-6.601-2.112-8.17l-1.304-.686c-2.987-1.569-7.962-1.902-11.113-.742z" />
      <path fill="#695445" stroke="#695445" stroke-width="5" d="M735.821 570.473l-2.761 9.352c-1.12 3.793-3.295 9.799-4.859 13.415l-1.912 4.422c-1.563 3.615-3.549 9.676-4.436 13.537l-2.091 9.107c-.887 3.861-1.31 10.195-.946 14.149l.954 10.364c.364 3.953 3.168 9.004 6.264 11.282l3.938 2.898c3.095 2.278 8.576 3.343 12.242 2.38l4.903-1.289c3.666-.963 8.744-4.081 11.343-6.965l4.222-4.685c2.599-2.883 6.122-8.075 7.87-11.596l2.76-5.56c1.747-3.521 4.176-9.415 5.424-13.164l3.042-9.135 4.653-14.673a414.246 414.246 0 004.043-13.737l2.3-8.463c1.038-3.819 1.376-10.09.755-14.008l-.023-.144c-.621-3.917-3.056-9.592-5.439-12.675l-1.217-1.575c-2.383-3.083-7.372-5.776-11.144-6.016l-11.336-.723c-3.771-.24-8.689 2.123-10.985 5.278l-11.38 15.644c-2.295 3.155-5.064 8.787-6.184 12.58z" />
      <path fill="#695445" stroke="#695445" stroke-width="5" d="M697.133 395.726l-4.154 1.378c-2.827.939-5.779 4.129-6.593 7.125l-1.595 5.868c-.814 2.997-.675 7.808.312 10.746l.971 2.893c.987 2.939 3.565 7.02 5.759 9.117l.387.37c2.194 2.096 4.591 6.236 5.355 9.247l.264 1.039a288.186 288.186 0 012.507 10.969l.023.111c.621 3.047 3.274 6.647 5.926 8.041l2.516 1.322c2.652 1.394 7.033 1.585 9.785.428l5.18-2.178c2.753-1.157 5.396-4.584 5.904-7.655l1.19-7.193c.508-3.07-.364-7.695-1.947-10.329l-3.356-5.586c-1.583-2.635-4.712-6.388-6.99-8.383l-.841-.737c-2.278-1.995-5.147-5.898-6.409-8.717l-3.774-8.429-2.901-6.099c-1.326-2.788-4.692-4.287-7.519-3.348z" />
      <path fill="#6a503e" stroke="#6a503e" stroke-width="3.25" d="M664.861 644.439l1.313 5.522c.683 2.872 2.629 7.103 4.346 9.449l1.356 1.853c1.717 2.346 3.796 6.537 4.645 9.36l1.473 4.904c.848 2.823 2.802 7.108 4.364 9.571l.402.634c1.562 2.463 5.043 5.042 7.776 5.76l2.224.584c2.732.718 7.126.585 9.814-.298l2.384-.783c2.688-.883 6.21-3.54 7.867-5.934l.06-.088c1.657-2.394 2.471-6.669 1.818-9.549l-1.424-6.29c-.653-2.88-3.028-6.626-5.306-8.368l-.083-.063c-2.278-1.741-6.306-3.859-8.998-4.729l-.098-.032c-2.692-.871-5.008-3.972-5.174-6.928l-.158-2.832c-.165-2.955-1.169-7.57-2.241-10.307l-1.419-3.622c-1.072-2.737-4.224-4.956-7.041-4.956h-6.464c-2.816 0-6.365 1.996-7.928 4.46l-1.916 3.022c-1.563 2.463-2.275 6.788-1.592 9.66z" />
      <path fill="#6a503e" stroke="#6a503e" stroke-width="3.25" d="M613.467 486.952l7.154-8.311c1.998-2.322 5.741-5.413 8.36-6.904l8.694-4.95c2.619-1.491 7.156-2.7 10.135-2.7h8.149c2.978 0 7.667.854 10.472 1.907l11.049 4.148c2.805 1.053 7.038 3.39 9.456 5.219l10.181 7.706c2.417 1.829 5.195 5.7 6.204 8.646l3.163 9.236c1.009 2.946.788 7.625-.492 10.452l-2.937 6.483c-1.28 2.826-4.478 6.253-7.142 7.653l-7.016 3.687c-2.664 1.4-7.216 2.192-10.167 1.769l-5.975-.856c-2.952-.423-6.793-2.796-8.58-5.301l-.345-.484c-1.787-2.504-5.598-5.058-8.513-5.703l-3.836-.848c-2.914-.645-7.667-1.527-10.616-1.97l-5.228-.785c-2.948-.443-7.472-1.989-10.105-3.454l-3.343-1.861c-2.632-1.464-5.846-4.922-7.178-7.723l-2.75-5.781c-1.332-2.801-.792-6.953 1.206-9.275z" />
      <path fill="#695445" stroke="#695445" stroke-width="5" d="M170.831 603.353l.036.168c.963 4.585 5.343 9.913 9.784 11.901l12.456 5.577c4.441 1.988 11.079 1.153 14.827-1.865l4.872-3.923c3.748-3.018 7.772-9.137 8.988-13.666l3.124-11.635c1.216-4.53 1.894-11.981 1.514-16.643l-1-12.27c-.38-4.662-4.47-9.615-9.135-11.063l-5.416-1.681c-4.666-1.448-10.599.567-13.253 4.5l-23.735 35.178c-2.654 3.933-4.024 10.838-3.062 15.422z" />
      <path fill="#6a503e" stroke="#6a503e" stroke-width="3.25" d="M241.079 650.195l-.098.182c-2.672 4.972-2.692 13.043-.044 18.026l1.093 2.057c2.648 4.983 9.433 10.007 15.156 11.221l.725.153c5.722 1.214 15.037 1.391 20.805.395l2.519-.435c5.769-.995 14.813-3.586 20.201-5.786l10.199-4.164c5.388-2.201 6.741-7.484 3.021-11.802l-4.975-5.775c-3.719-4.318-11.378-6.853-17.106-5.663l-1.416.294c-5.727 1.189-14.926.866-20.546-.723l-14.52-4.105c-5.62-1.589-12.342 1.154-15.014 6.125z" />
      <path fill="#695445" stroke="#695445" stroke-width="2.06" d="M285.533 601.722l2.531.752c6.313 1.873 16.766 2.894 23.347 2.279l6.145-.573c6.58-.615 14.577-5.546 17.861-11.015l4.963-8.266c3.283-5.47 7.431-14.811 9.264-20.865l.956-3.156c1.833-6.054.418-15.257-3.16-20.555l-.131-.194c-3.579-5.299-11.835-9.374-18.442-9.102l-.497.02c-6.606.272-16.544 3.142-22.196 6.412l-2.648 1.531c-5.652 3.269-13.729 9.793-18.04 14.57l-.678.751c-4.311 4.778-12.358 11.347-17.974 14.673l-5.244 3.106c-5.616 3.326-11.168 11.041-12.401 17.231l-.352 1.765c-1.233 6.191.669 15.504 4.248 20.802l.131.194c3.578 5.299 9.196 5.191 12.548-.241l2.269-3.676c3.351-5.432 11.186-8.316 17.5-6.443z" />
      <path fill="#ff4e47" stroke="#ff4e47" stroke-width="28" d="M 707.78417,358.59878 846.34934,503.2334" />
    </g>
    <g class="head">
      <path fill="#bd8b4a" d="M911.034 474.462l58.711 25.635c3.998 1.746 10.859 2.947 15.323 2.683l8.328-.493c4.465-.264 10.914-2.315 14.884-4.112l7.47-3.676c3.97-1.797 10.92-4.795 14.4-7.235l14.84-10.392c3.48-2.441 8.1-7.154 10.33-10.528l7.47-11.338c2.22-3.374 3.57-6.514 3.01-7.015 0 0-89.032-79.584-89.326-80.524" />
      <path fill="#bd8b4a" d="M913.883 547.305C869.89 505.578 853.078 442.761 876.334 407c23.255-35.762 77.772-30.925 121.766 10.802 43.99 41.727 62.76 40.306 62.76 40.306-23.26 35.762-102.983 130.924-146.977 89.197z" />
      <path fill="#bd8b4a" d="m 878.57469,361.46519 c -3.549,8.006 18.32731,62.81481 18.32731,62.81481 0,0 -5.576,45.517 7.976,50.611 1.64,0.617 9.306,-0.397 19.932,-4.902 30.681,-13.007 85.9,-42.08 90.94,-42.892 3.24,-0.523 -36.398,-6.798 -33.634,-8.579 11.678,-7.525 16.195,-66.177 16.195,-66.177 0,0 0.84,-18.025 -3.737,-33.088 -3.332,-10.962 -11.307,-21.389 -20.966,-26.835 -4.181,-2.358 -8.527,-2.554 -13.033,-3.118 -9.386,-1.176 -17.697,-0.873 -26.857,2.227 -20.387,6.897 -64.15231,47.92519 -55.14331,69.93819" />
      <g class="eye">
        <ellipse ry="58.79958" rx="63.385227" cy="337.63333" cx="930.31775" opacity="1" fill="#6f5a4b" fill-opacity="1" stroke="#6f5a4b" stroke-width="7" />
        <path fill="#fff" stroke="#fff" stroke-width="3.25" d="M917.372 339.256c0-19.15 15.843-34.674 35.388-34.674 19.544 0 35.388 15.524 35.388 34.674 0 19.149-15.844 34.673-35.388 34.673-12.483 0-23.457-6.332-29.757-15.898a34.015 34.015 0 01-5.631-18.775z" />
        <path fill="#262626" stroke="#262626" d="M953.463 334.106c0-8.041 6.456-14.56 14.421-14.56 7.965 0 14.421 6.519 14.421 14.56s-6.456 14.559-14.421 14.559c-7.965 0-14.421-6.518-14.421-14.559z" />
        <path fill="#fff" stroke="#fff" stroke-width="5" d="M972.119 327.418c0-.661.711-1.197 1.589-1.197.877 0 1.589.536 1.589 1.197s-.712 1.197-1.589 1.197c-.878 0-1.589-.536-1.589-1.197z" />
        <g class="closed-eye" fill="#6a503e" stroke="#4a392c" stroke-width="3">
          <path d="M917.003 339.068c0-19.357 16.007-35.048 35.754-35.048 19.746 0 35.753 15.691 35.753 35.048 0 19.356-16.007 35.048-35.753 35.048-12.612 0-23.699-6.401-30.064-16.07a34.386 34.386 0 01-5.69-18.978z" />
          <path d="M923.322 358.879c.132-4.19 10.549-11.275 13.553-13.265 11.34-7.511 24.072-11.238 37.223-13.485 4.288-.732 8.655-.216 12.924-.697" />
        </g>
      </g>
      <g class="nose">
        <path fill="#4a392c" d="M1035.88 441.045c5.38-8.889 16.28-12.745 24.36-8.614 8.08 4.131 10.27 14.686 4.89 23.574-5.38 8.889-16.29 12.746-24.37 8.615-8.07-4.131-10.26-14.686-4.88-23.575z" />
        <path fill="#fff" stroke="#fff" stroke-width="1.74" d="M1051.34 441.56c.03 2.329 2.27-.018 2.64-1.768.18-.862-1.8 2-.93 2.516 1.16.684 2.36-2.411 1.42-3.282-2.95-2.722-2.72 7.429-1.07 3.998 1.03-2.155-5.45 3.379-3.62-2.029 2.85-8.382 9.1 2.281 2.45 3.226-4.55.645-3.63-8.619 2.7-6.739 4.48 1.33-.01 5.966-3.4 5.448-5.35-.818.82-8.238 3.9-4.094" opacity=".965" />
      </g>
      <g>
        <path class="ball" fill="#bdf971" stroke="#bdf971" d="m 930.80242,477.19065 c -13.05851,3.00801 -24.77881,15.97694 -26.81485,30.67673 -2.03603,14.69979 5.61282,31.12503 15.54844,43.47271 9.93563,12.34768 22.15522,20.61485 36.12779,24.29677 13.97258,3.68192 29.69644,2.77869 42.91157,-2.81649 13.21513,-5.59518 23.92113,-15.88139 30.65943,-28.14324 6.7383,-12.26185 9.5079,-26.49697 5.7369,-38.5172 -3.771,-12.02022 -14.0825,-21.82316 -25.5603,-24.46605 -11.47786,-2.6429 -24.12011,1.87601 -37.63861,0.6591 -13.5185,-1.21691 -27.91187,-8.17033 -40.97037,-5.16233 z" />
        <clipPath id="ballClip">
          <path class="ball" fill="#bdf971" stroke="#bdf971" d="m 930.80242,477.19065 c -13.05851,3.00801 -24.77881,15.97694 -26.81485,30.67673 -2.03603,14.69979 5.61282,31.12503 15.54844,43.47271 9.93563,12.34768 22.15522,20.61485 36.12779,24.29677 13.97258,3.68192 29.69644,2.77869 42.91157,-2.81649 13.21513,-5.59518 23.92113,-15.88139 30.65943,-28.14324 6.7383,-12.26185 9.5079,-26.49697 5.7369,-38.5172 -3.771,-12.02022 -14.0825,-21.82316 -25.5603,-24.46605 -11.47786,-2.6429 -24.12011,1.87601 -37.63861,0.6591 -13.5185,-1.21691 -27.91187,-8.17033 -40.97037,-5.16233 z" />
        </clipPath>
        <path class="ball-decoration" clip-path="url(#ballClip)" fill="none" stroke="#fff" stroke-width="4" d="m 963.39546,597.71943 c 21.49913,-19.30313 4.80913,-64.80408 16.71919,-83.46282 16.19467,-25.37116 67.93925,-22.92156 89.43095,-47.68524" />
        <g class="ball-sound">
          <path fill="none" stroke="#B5D8E8" stroke-width="4" d="m 1096.3827,564.72204 c 11.3886,12.59779 6.8073,-1.45406 10.931,-3.9212 1.9116,-1.14344 5.2815,9.1526 9.7911,16.39495 0.2794,0.44804 0.3148,-1.06074 0.4718,-1.59122 0.9548,-3.21847 1.6781,-12.7846 4.2464,-14.32081 2.4235,-1.4498 4.5316,8.986 8.6637,9.21558 5.7286,0.31843 3.6488,-6.24515 5.8457,-8.73267 1.4614,-1.65478 3.9233,1.78478 6.0289,1.62906 4.8167,-0.35637 4.1527,-6.99656 8.0084,0.44507" />
          <path fill="none" stroke="#B5D8E8" stroke-width="4" d="m 1070.9116,599.79298 c -2.9954,2.45576 -9.6186,6.11285 -7.2109,10.56543 0.5391,0.99715 13.1337,-2.24375 13.7587,-2.44022 0.5954,-0.18689 1.7073,-1.42002 1.6666,-0.90122 -0.083,1.06689 -3.3427,7.9055 -1.5136,9.15131 2.703,1.84098 10.413,-0.47946 10.1532,2.84261 -0.344,4.40226 -7.7885,9.49729 -5.8163,13.14464 1.3149,2.43155 6.6186,0.47256 9.456,1.55305 4.0577,1.54512 -3.7339,7.61167 2.5171,8.63834" />
        </g>
      </g>
      <path fill="#695445" stroke="#695445" stroke-width="3.25" d="M894.558 290.191l5.447-7.287c.796-1.066 2.311-2.539 3.384-3.291l5.204-3.647c1.073-.751 2.934-1.708 4.157-2.137l6.932-2.428c1.223-.429 3.224-.493 4.47-.144l6.85 1.92c1.246.349 2.667 1.642 3.175 2.888l.433 1.062c.508 1.245.24 3.089-.599 4.118l-1.506 1.847c-.839 1.029-2.454 2.354-3.608 2.961l-8.515 4.475-3.971 2.087c-1.154.606-2.741 1.956-3.547 3.014l-4.601 6.045-3.025 3.71c-.839 1.029-2.558 1.984-3.839 2.134l-2.176.254c-1.282.15-3.06-.506-3.971-1.464l-.486-.51c-.912-.959-2.018-2.764-2.47-4.033l-.633-1.774c-.453-1.269-.408-3.307.1-4.553l.433-1.062c.508-1.246 1.565-3.119 2.362-4.185z" />
      <g class="ear">
        <path fill="#554132" d="M899.667 340.093c-18.458-82.741-41.446-100.776-75.085-74.525-35.063 27.363-93.432 28.614-74.975 111.355 18.457 82.741 67.012 141.572 108.449 131.401 41.438-10.17 60.068-85.49 41.611-168.231z" />
        <path fill="#4a392c" stroke="#4a392c" stroke-width="3.55" d="M771.398 397.916l-.417 2.678c-.344 2.211.02 5.69.813 7.77l.614 1.609c.793 2.08 2.61 5.099 4.059 6.743l1.036 1.176a113.619 113.619 0 005.474 5.712l3.036 2.913c1.575 1.511 3.888 4.189 5.167 5.981l.047.066c1.279 1.793 3.799 4.178 5.63 5.327l1.699 1.067c1.83 1.149 4.911 2.774 6.881 3.629l6.158 2.674 6.859 2.695c1.997.784 5.282 1.902 7.338 2.497l3.679 1.064c2.056.595 5.445 1.213 7.571 1.38l6.965.549 9.208.371c2.13.086 5.581.038 7.709-.108l2.904-.199c2.128-.145 5.485-.858 7.499-1.592l3.653-1.331c2.014-.734 5.05-2.389 6.781-3.697l5.562-4.202c1.731-1.308 3.794-4.047 4.608-6.118l1.07-2.723c.814-2.071 1.239-5.549.949-7.769l-.289-2.215c-.29-2.22-1.202-5.691-2.037-7.753l-.359-.885c-.835-2.062-2.629-5.119-4.008-6.829l-3.889-4.823-3.114-4.076c-1.335-1.747-4.006-3.877-5.966-4.758l-1.023-.46c-1.96-.881-5.239-1.978-7.323-2.451l-4.892-1.109-4.17-1.093a107.68 107.68 0 01-7.401-2.285l-3.915-1.398a183.715 183.715 0 01-7.245-2.794l-5.752-2.388c-1.983-.823-5.273-1.906-7.348-2.418l-2.754-.681a176.174 176.174 0 01-7.466-2.057l-10.622-3.234-5.07-1.592c-2.043-.642-5.417-1.359-7.536-1.602l-1.639-.188c-2.119-.243-4.799 1.069-5.986 2.931l-.779 1.223c-1.187 1.861-2.653 5.108-3.275 7.251l-.912 3.142c-.622 2.143-2.333 5.18-3.822 6.784l-.64.688c-1.489 1.604-2.976 4.696-3.32 6.908z" />
        <path fill="#6a503e" stroke="#6a503e" stroke-width="3.25" d="M818.613 319.578l-3.003 6.664c-.639 1.418-1.214 3.839-1.284 5.407l-.428 9.507-.135 10.621.142 9.336c.024 1.569.501 4.019 1.066 5.472l2.902 7.46c.565 1.453 1.687 3.695 2.506 5.007l2.82 4.52c.819 1.312 2.465 3.12 3.676 4.038l3.88 2.939c1.211.918 3.402 1.601 4.893 1.525l3.803-.193c1.491-.076 3.62-.964 4.755-1.984l4.39-3.945c1.135-1.02 2.43-3.056 2.892-4.549l1.583-5.112c.462-1.493 1.038-3.957 1.287-5.505l1.136-7.078c.248-1.548.456-4.075.464-5.644l.045-8.997c.008-1.569-.242-4.085-.558-5.619l-1.903-9.229c-.316-1.534-1.056-3.944-1.652-5.383l-2.468-5.954c-.596-1.439-1.866-3.573-2.836-4.766l-3.753-4.614c-.97-1.193-2.94-2.43-4.399-2.762l-3.66-.833c-1.459-.332-3.767-.131-5.154.449l-3.538 1.479c-1.387.58-3.354 1.964-4.393 3.092l-.038.041c-1.039 1.127-2.399 3.191-3.038 4.61z" />
        <path fill="#695445" stroke="#695445" stroke-width="2.06" d="M828.818 330.254l-.037.064c-1.036 1.741-2.292 4.758-2.806 6.739l-.411 1.584c-.514 1.981-1.088 5.243-1.282 7.286l-.054.563c-.194 2.043.417 5.154 1.366 6.949l.352.666c.948 1.795 3.275 2.952 5.197 2.584l1.372-.262c1.922-.367 4.493-1.943 5.744-3.521l3.045-3.841c1.251-1.578 2.908-4.378 3.701-6.254l.157-.37c.793-1.877 1.747-5.03 2.13-7.043l.127-.671c.384-2.013-.603-4.6-2.203-5.777l-1.779-1.309c-1.601-1.178-4.463-1.879-6.394-1.567l-2.854.462c-1.931.312-4.335 1.977-5.371 3.718z" />
      </g>
    </g>
  </g>
</svg>

<style>
.sausage-dog-animation {
  height: 15rem;
}

.ear,
.closed-eye,
.lean,
.front-legs,
.leg,
.head,
.tail,
.tail-blur,
.shadow {
  animation-duration: 3s;
  animation-iteration-count: infinite;
  animation-timing-function: linear;
}

.ball,
.ball-decoration,
.ball-sound {
  animation-duration: 3s;
  animation-iteration-count: infinite;
  animation-timing-function: ease-in-out;
}

.ball {
  animation-name: squishBall;
  transform: matrix(
    1.0951654,
    0.52195853,
    -0.52866476,
    1.2371611,
    208.27138,
    -632.28196
  );
}

@keyframes squishBall {
  0%,
  50%,
  72%,
  80%,
  92%,
  100% {
    d: path(
      "m 930.80242,477.19065 c -13.05851,3.00801 -24.77881,15.97694 -26.81485,30.67673 -2.03603,14.69979 5.61282,31.12503 15.54844,43.47271 9.93563,12.34768 22.15522,20.61485 36.12779,24.29677 13.97258,3.68192 29.69644,2.77869 42.91157,-2.81649 13.21513,-5.59518 23.92113,-15.88139 30.65943,-28.14324 6.7383,-12.26185 9.5079,-26.49697 5.7369,-38.5172 -3.771,-12.02022 -14.0825,-21.82316 -25.5603,-24.46605 -11.47786,-2.6429 -24.12011,1.87601 -37.63861,0.6591 -13.5185,-1.21691 -27.91187,-8.17033 -40.97037,-5.16233 z"
    );
    transform: matrix(
      1.0951654,
      0.52195853,
      -0.52866476,
      1.2371611,
      208.27138,
      -632.28196
    );
  }
  65%,
  85% {
    d: path(
      "m 932.4158,479.26229 c -14.67189,0.93637 -26.39219,13.9053 -28.42823,28.60509 -2.03603,14.69979 5.61282,31.12503 15.94228,38.97676 10.32947,7.85173 23.3365,7.12856 34.39281,10.90173 11.0563,3.77318 20.16082,12.04166 33.72906,12.76579 13.56828,0.72414 31.59878,-6.09489 42.72528,-18.16177 11.1265,-12.06689 15.3474,-29.37891 10.9229,-43.93943 -4.4245,-14.56052 -17.4943,-26.36654 -28.8312,-27.22343 -11.337,-0.85689 -20.93852,9.23684 -34.54977,8.73602 -13.61125,-0.50082 -31.23124,-11.59712 -45.90313,-10.66076 z"
    );
    transform: matrix(
      1.0951654,
      0.52195853,
      -0.52866476,
      1.2371611,
      208.27138,
      -642.28196
    );
  }
}

.ball-decoration {
  animation-name: ballDecorationAnimation;
}

@keyframes ballDecorationAnimation {
  0%,
  50%,
  72%,
  80%,
  92%,
  100% {
    d: path(
      "m 963.39546,597.71943 c 21.49913,-19.30313 4.80913,-64.80408 16.71919,-83.46282 16.19467,-25.37116 67.93925,-22.92156 89.43095,-47.68524"
    );
  }
  65%,
  85% {
    d: path(
      "m 978.40243,581.77452 c 21.49916,-19.30313 -15.82546,-51.20401 -3.9154,-69.86275 16.19467,-25.37116 63.71847,-19.16982 85.21017,-43.9335"
    );
  }
}

.ball-sound {
  animation-name: ballSound;
  visibility: hidden;
}

@keyframes ballSound {
  0%,
  60%,
  70%,
  80%,
  90%,
  100% {
    visibility: hidden;
    transform: translateY(0);
  }
  65%,
  67%,
  69%,
  85%,
  87%,
  89% {
    visibility: visible;
    transform: translateY(-3px);
  }
  66%,
  68%,
  86%,
  88% {
    visibility: visible;
    transform: translateY(3px);
  }
}

.ear {
  animation-name: moveEar;
  transform-origin: top center;
  transform-box: fill-box;
}

@keyframes moveEar {
  0%,
  12%,
  21%,
  31%,
  35%,
  100% {
    transform: rotateZ(0);
  }
  9%,
  19%,
  29% {
    transform: rotateZ(-5deg);
    transform: rotateZ(-10deg);
  }
  13%,
  23%,
  33% {
    transform: rotateZ(5deg);
    transform: rotateZ(10deg);
  }
}

.closed-eye {
  animation-name: closeEye;
}

@keyframes closeEye {
  0%,
  50%,
  100% {
    visibility: hidden;
  }
  10% {
    visibility: visible;
  }
}

.lean {
  animation-name: leanDown;
  transform-origin: center;
}

@keyframes leanDown {
  0%,
  50%,
  100% {
    transform: rotateZ(0) translateY(0);
  }
  60%,
  90% {
    transform: rotateZ(10deg) translateY(5%);
  }
}

.front-legs {
  animation-name: flexLegs;
}

@keyframes flexLegs {
  0%,
  50%,
  100% {
    transform: translateX(0);
  }
  60%,
  90% {
    transform: translateX(12%);
  }
}

.leg {
  animation-name: rotateLegs;
  transform-origin: bottom left;
  transform-box: fill-box;
  transform: translateX(16%) rotate(-10deg);
}

@keyframes rotateLegs {
  0%,
  50%,
  100% {
    transform: translateX(16%) rotate(-10deg);
  }
  60%,
  90% {
    transform: translateX(35%) rotate(-83deg);
  }
}

.head {
  animation-name: lookDown;
  transform-origin: top right;
  transform-box: fill-box;
}

@keyframes lookDown {
  0%,
  55%,
  100% {
    transform: rotateZ(0) translate(0, 0);
  }
  60%,
  90% {
    transform: rotateZ(5deg) translate(2.5%, 6%);
  }
}

.tail {
  animation-name: moveTail;
  transform-origin: bottom center;
}

@keyframes moveTail {
  0%,
  50%,
  90%,
  100% {
    d: path(
      "m 161.6285,568.63016 20.92664,-20.00034 C 151.50961,521.73829 14.092,504.432 14.092,504.432 c 0,0 128.8135,26.71916 147.5365,64.19816 z"
    );
  }
  64%,
  70%,
  76%,
  82% {
    d: path(
      "m 161.6285,568.63016 20.92664,-20.00034 C 151.50961,521.73829 77.565044,422.94078 77.565044,422.94078 c 0,0 65.340456,108.21038 84.063456,145.68938 z"
    );
  }
  60%,
  66%,
  72%,
  78%,
  84% {
    d: path(
      "m 161.6285,568.63016 20.92664,-20.00034 C 151.50961,521.73829 14.092,504.432 14.092,504.432 c 0,0 128.8135,26.71916 147.5365,64.19816 z"
    );
  }
  62%,
  68%,
  74%,
  80%,
  86% {
    d: path(
      "m 161.6285,568.63016 20.92664,-20.00034 C 151.50961,521.73829 159.419,390.74 159.419,390.74 c 0,0 -16.5135,140.41116 2.2095,177.89016 z"
    );
  }
}

.tail-blur {
  animation-name: tailBlur;
  transform-origin: bottom center;
}

@keyframes tailBlur {
  0%,
  59%,
  90%,
  100% {
    opacity: 0;
  }
  64%,
  70%,
  76%,
  82% {
    transform: rotate(-2deg);
    opacity: 0;
  }
  60%,
  66%,
  72%,
  78%,
  84% {
    opacity: 0.3;
  }
  62%,
  68%,
  74%,
  80%,
  86% {
    transform: rotate(2deg);
    opacity: 0;
  }
}

.shadow {
  animation-name: scaleShadow;
  transform-origin: center center;
}

@keyframes scaleShadow {
  0%,
  55%,
  100% {
    transform: scaleX(1) translateX(0);
  }
  60%,
  90% {
    transform: scaleX(1.1) translateX(4%);
  }
}
</style>
`;

  const loader11 = `
<!-- https://codepen.io/JayJay89/pen/aNmoYR -->
			<div class="corgi">

				<div class="head">
					<div class="ear ear--r"></div>
					<div class="ear ear--l"></div>

					<div class="eye eye--left"></div>
					<div class="eye eye--right"></div>

					<div class="face">
						<div class="face__white">
							<div class=" face__orange face__orange--l"></div>
							<div class=" face__orange face__orange--r"></div>
						</div>
					</div>

					<div class="face__curve"></div>

					<div class="mouth">

						<div class="nose"></div>
						<div class="mouth__left">
							<div class="mouth__left--round"></div>
							<div class="mouth__left--sharp"></div>
						</div>
						
						<div class="lowerjaw">
							<div class="lips"></div>
							<div class="tongue test"></div>
						</div>

						<div class="snout"></div>
					</div>
				</div>
				
				<div class="neck__back"></div>
				<div class="neck__front"></div>

				<div class="body">
					<div class="body__chest"></div>
				</div>

				<div class="foot foot__left foot__front foot__1"></div>
				<div class="foot foot__right foot__front foot__2"></div>
				<div class="foot foot__left foot__back foot__3"></div>
				<div class="foot foot__right foot__back foot__4"></div>

				<div class="tail test"></div>
			</div>

      <style>
      /*z-indices*/
/*Animation*/
@keyframes eye-blink {
  0% {
    transform: scaleY(1);
  }
  10% {
    transform: scaleY(0.1);
  }
  20% {
    transform: scaleY(1);
  }
  100% {
    transform: scaleY(1);
  }
}
@keyframes tail-wag {
  0% {
    transform: rotate(-25deg);
  }
  10% {
    transform: rotate(0deg);
  }
  20% {
    transform: rotate(-25deg);
  }
  30% {
    transform: rotate(0deg);
  }
  40% {
    transform: rotate(-25deg);
  }
  50% {
    transform: rotate(0deg);
  }
  60% {
    transform: rotate(-25deg);
  }
  70% {
    transform: rotate(0deg);
  }
  100% {
    transform: rotate(-25deg);
  }
}
@keyframes tongue-stick {
  0% {
    transform: scaleY(0.1) translateY(-20px);
  }
  20% {
    transform: scaleY(0.1) translateY(-20px);
  }
  30% {
    transform: scaleY(0.5) translateY(0px);
  }
  40% {
    transform: scaleY(1) translateY(0px) rotate(0deg);
  }
  50% {
    transform: scaleY(0.8) translateY(0px) rotate(15deg);
  }
  60% {
    transform: scaleY(1) translateY(0px) rotate(0deg);
  }
  70% {
    transform: scaleY(0.8) translateY(0px) rotate(-15deg);
  }
  80% {
    transform: scaleY(1) translateY(0px) rotate(0deg);
  }
  90% {
    transform: scaleY(0.8) translateY(0px) rotate(15deg);
  }
  100% {
    transform: scaleY(0.1) translateY(-20px) rotate(0deg);
  }
}
@keyframes ear-shake-right {
  0% {
    transform: rotate(70deg);
  }
  10% {
    transform: rotate(80deg);
  }
  30% {
    transform: rotate(70deg);
  }
  40% {
    transform: rotate(80deg);
  }
  100% {
    transform: rotate(70deg);
  }
}
@keyframes ear-shake-left {
  0% {
    transform: rotate(-70deg);
  }
  10% {
    transform: rotate(-80deg);
  }
  30% {
    transform: rotate(-70deg);
  }
  40% {
    transform: rotate(-80deg);
  }
  100% {
    transform: rotate(-70deg);
  }
}
@keyframes body-shake {
  0% {
    transform: translateY(0px);
  }
  16.6666666667% {
    transform: translateY(2%);
  }
  33.3333333333% {
    transform: translateY(0px);
  }
  50% {
    transform: translateY(2%);
  }
  66.6666666667% {
    transform: translateY(0px);
  }
  83.3333333333% {
    transform: translateY(2%);
  }
  100% {
    transform: translateY(0px);
  }
}
@keyframes paw-press {
  0% {
    transform: scaleY(1) scaleX(1);
  }
  16.6666666667% {
    transform: scaleY(0.9) scaleX(1.05) translateY(10%);
  }
  33.3333333333% {
    transform: scaleY(1) scaleX(1);
  }
  50% {
    transform: scaleY(0.9) scaleX(1.05) translateY(10%);
  }
  66.6666666667% {
    transform: scaleY(1) scaleX(1);
  }
  83.3333333333% {
    transform: scaleY(0.9) scaleX(1.05) translateY(10%);
  }
  100% {
    transform: scaleY(1) scaleX(1);
  }
}
@keyframes neck-shake {
  0% {
    top: 9%;
  }
  16.6666666667% {
    top: 11%;
  }
  33.3333333333% {
    top: 9%;
  }
  50% {
    top: 11%;
  }
  66.6666666667% {
    top: 9%;
  }
  83.3333333333% {
    top: 11%;
  }
  100% {
    top: 9%;
  }
}
@keyframes head-shake {
  0% {
    top: 6%;
  }
  16.6666666667% {
    top: 8%;
  }
  33.3333333333% {
    top: 6%;
  }
  50% {
    top: 8%;
  }
  66.6666666667% {
    top: 6%;
  }
  83.3333333333% {
    top: 8%;
  }
  100% {
    top: 6%;
  }
}
@keyframes mouth-shake {
  0% {
    bottom: 0%;
  }
  16.6666666667% {
    bottom: 2%;
  }
  33.3333333333% {
    bottom: 0%;
  }
  50% {
    bottom: 2%;
  }
  66.6666666667% {
    bottom: 0%;
  }
  83.3333333333% {
    bottom: 2%;
  }
  100% {
    bottom: 0%;
  }
}

#rotate {
  background-color: #f0f0f0;
  padding: 5px;
  position: fixed;
  top: 0px;
  left: 50px;
}

.corgi {
  height: 180px;
  width: 252px;
  position: relative;
}
.corgi div {
  position: absolute;
}
.corgi .ear {
  background-color: #F09F2E;
  height: 30%;
  width: 55%;
  top: 5%;
  z-index: 3;
}
.corgi .ear--r {
  right: 75%;
  border-bottom-left-radius: 100% 90%;
  border-top-left-radius: 10%;
  transform-origin: 80% center;
  animation: ear-shake-right 2s none infinite;
}
.corgi .ear--l {
  left: 63%;
  background-color: #D27537;
  border-bottom-right-radius: 100% 90%;
  border-top-right-radius: 10%;
  transform-origin: 20% center;
  animation: ear-shake-left 2s none infinite;
}
.corgi .head {
  top: 6%;
  right: 10%;
  height: 40%;
  width: 30%;
  z-index: 3;
  animation: head-shake 2s none infinite;
}
.corgi .face {
  background-color: #F09F2E;
  border-radius: 50%;
  overflow: hidden;
  height: 100%;
  width: 100%;
  z-index: 4;
}
.corgi .eye {
  background-color: #3E3954;
  height: 6%;
  width: 6%;
  position: absolute;
  z-index: 6;
  border-radius: 50%;
  animation: eye-blink 2s none infinite;
}
.corgi .eye--left {
  left: 40%;
  top: 43%;
}
.corgi .eye--right {
  right: 13%;
  top: 41%;
}
.corgi .face__white {
  background-color: #FFFFFF;
  width: 45%;
  height: 77%;
  top: -15%;
  left: 29%;
  transform: rotate(-25deg);
}
.corgi .face__orange {
  background-color: #F09F2E;
  content: " ";
  position: absolute;
  width: 110%;
  height: 110%;
  display: block;
  border-radius: 100%;
}
.corgi .face__orange--l {
  right: 65%;
}
.corgi .face__orange--r {
  left: 65%;
}
.corgi .face__curve {
  background-color: #F09F2E;
  width: 30%;
  height: 20%;
  right: -12%;
  bottom: 42%;
  overflow: hidden;
}
.corgi .face__curve:after {
  content: "";
  background-color: #8C5A46;
  position: absolute;
  width: 69%;
  height: 82%;
  border-radius: 0% 100%;
  top: -32%;
  right: -13%;
}
.corgi .mouth {
  bottom: 0%;
  width: 100%;
  height: 50%;
  left: 28%;
  z-index: 5;
  animation: mouth-shake 2s none infinite;
}
.corgi .nose {
  height: 36%;
  width: 27%;
  top: 0%;
  background-color: #3E3954;
  z-index: 1;
  right: 0%;
  border-bottom-right-radius: 50% 100%;
  border-bottom-left-radius: 50% 100%;
}
.corgi .nose:after {
  content: "";
  width: 100%;
  height: 30%;
  display: block;
  border-top-right-radius: 50% 100%;
  border-top-left-radius: 50% 100%;
  background-color: #3E3954;
  position: absolute;
  top: -25%;
}
.corgi .mouth__left {
  background-color: #FFFFFF;
  width: 50%;
  height: 55%;
}
.corgi .mouth__left--round {
  background-color: #F09F2E;
  width: 100%;
  height: 100%;
  border-radius: 100%;
  left: -50%;
  top: -50%;
}
.corgi .mouth__left--sharp {
  background-color: #F09F2E;
  width: 35%;
  height: 50%;
  bottom: 0px;
  left: -20%;
  transform: skewX(50deg);
}
.corgi .lowerjaw {
  background-color: #FFFFFF;
  width: 100%;
  height: 80%;
  border-radius: 50%/100%;
  border-top-left-radius: 0;
  border-top-right-radius: 0;
  bottom: -9%;
}
.corgi .lips {
  z-index: 2;
  height: 25%;
  width: 35%;
  top: 19%;
  right: 2%;
}
.corgi .lips:before, .corgi .lips:after {
  content: "";
  display: block;
  background: #FFFFFF;
  width: 39%;
  height: 100%;
  border-color: #3E3954;
  border-width: 3px;
  border-style: solid;
  border-bottom-right-radius: 50%;
  border-bottom-left-radius: 50%;
  border-top-left-radius: 40%;
  border-top-right-radius: 20%;
  border-top-color: transparent;
  position: absolute;
}
.corgi .lips:before {
  z-index: 1;
}
.corgi .lips:after {
  transform: rotateY(180deg);
  left: initial;
  right: 9%;
}
.corgi .tongue {
  width: 15%;
  height: 60%;
  background-color: #F15F55;
  right: 14%;
  top: 35%;
  border-bottom-right-radius: 50% 50%;
  border-bottom-left-radius: 50% 50%;
  transform-origin: 50% 0%;
  animation: tongue-stick 2s none infinite;
}
.corgi .snout {
  background-color: #FFFFFF;
  right: 0%;
  top: 0%;
  width: 50%;
  height: 36%;
  border-top-right-radius: 35% 75%;
}
.corgi .neck__back {
  height: 50%;
  width: 20%;
  transform: skewX(-20deg);
  background-color: #F09F2E;
  z-index: 2;
  right: 24%;
  top: 9%;
  animation: neck-shake 2s none infinite;
}
.corgi .neck__front {
  height: 50%;
  width: 20%;
  right: 11%;
  top: 20%;
  background-color: #F09F2E;
  z-index: 2;
  transform: skewX(2deg);
}
.corgi .body {
  height: 44%;
  width: 77%;
  background-color: #F09F2E;
  right: 10.5%;
  bottom: 12%;
  border-top-left-radius: 20% 50%;
  border-bottom-left-radius: 20% 50%;
  border-top-right-radius: 20% 60%;
  border-bottom-right-radius: 20% 40%;
  z-index: 2;
  overflow: hidden;
  animation: body-shake 2s none infinite;
}
.corgi .body__chest {
  background-color: #FFFFFF;
  height: 87%;
  width: 29%;
  right: 5%;
  bottom: -3%;
  border-top-left-radius: 50% 40%;
  border-top-right-radius: 50% 40%;
}
.corgi .foot {
  height: 35%;
  width: 9.5%;
  bottom: 0;
}
.corgi .foot__left {
  z-index: 3;
  background-color: #F09F2E;
}
.corgi .foot__left:after {
  background-color: #FFFFFF;
}
.corgi .foot__left:before {
  background-color: #F09F2E;
}
.corgi .foot__right {
  z-index: 1;
  background-color: #D27537;
}
.corgi .foot__right:after {
  background-color: #B6D8EF;
}
.corgi .foot__right:before {
  background-color: #D27537;
}
.corgi .foot__back:before {
  transform: skewX(-10deg);
  right: -25%;
}
.corgi .foot__front:before {
  transform: skewX(10deg);
  right: 25%;
}
.corgi .foot__1 {
  right: 37%;
}
.corgi .foot__2 {
  right: 15%;
}
.corgi .foot__2:before {
  transform: skewX(-10deg);
  right: -25%;
}
.corgi .foot__3 {
  left: 12.65%;
}
.corgi .foot__4 {
  left: 31%;
}
.corgi .foot:before {
  content: "";
  position: absolute;
  height: 100%;
  width: 100%;
  display: block;
}
.corgi .foot:after {
  /*paws*/
  content: "";
  position: absolute;
  bottom: 0;
  left: 0;
  width: 125%;
  height: 18%;
  border-top-left-radius: 50% 100%;
  border-top-right-radius: 50% 100%;
  animation: paw-press 2s none infinite;
}
.corgi .tail {
  width: 26%;
  height: 13%;
  background-color: #D27537;
  border-top-left-radius: 50% 100%;
  border-bottom-left-radius: 50% 100%;
  border-top-right-radius: 50% 100%;
  border-bottom-right-radius: 50% 100%;
  bottom: 40%;
  left: 1%;
  transform-origin: 80% center;
  animation: tail-wag 2s none infinite;
}

.test, .testrev {
  transition: all 0.8s ease;
}
  </style>
`

  const loader12 = `
<!-- From Uiverse.io by https://uiverse.io/chase2k25/brown-dog-100 --> 
<div class="loader-chase2k25">
  <div class="cube">
    <div class="face front">
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
    </div>
    <div class="face back">
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
    </div>
    <div class="face right">
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
    </div>
    <div class="face left">
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
    </div>
    <div class="face top">
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
    </div>
    <div class="face bottom">
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
      <div class="sticker"></div>
    </div>
  </div>
</div>
<style>
.loader-chase2k25 {
  position: absolute;
  bottom: 50px;
  left: 50px;
  width: 180px;
  height: 180px;
  transform-style: preserve-3d;
  animation: rotate-cube 8s cubic-bezier(0.76, 0, 0.24, 1) infinite;
  transition: transform 0.6s cubic-bezier(0.68, -0.55, 0.27, 1.55);
}

.loader:hover {
  animation-play-state: paused;
  transform: scale(1.15); /* Scaling only on hover */
}

.cube {
  position: absolute;
  width: 100%;
  height: 100%;
  transform-style: preserve-3d;
  transform: rotateX(35.264deg) rotateY(45deg); /* Isometric angle for square consistency */
}

.face {
  position: absolute;
  width: 180px;
  height: 180px;
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  grid-gap: 6px;
  padding: 6px;
  box-sizing: border-box;
  background: rgba(10, 20, 30, 0.9);
  border: 3px solid rgba(255, 255, 255, 0.15);
  transition: box-shadow 0.4s cubic-bezier(0.4, 0, 0.6, 1);
}

.face:hover {
  box-shadow: 0 0 25px rgba(255, 255, 255, 0.4);
}

.sticker {
  background: #fff;
  border-radius: 8px;
  animation: sticker-bounce 1.5s cubic-bezier(0.68, -0.55, 0.27, 1.55) infinite;
}

/* Face positions */
.front {
  transform: translateZ(90px);
}
.back {
  transform: translateZ(-90px) rotateY(180deg);
}
.right {
  transform: translateX(90px) rotateY(90deg);
}
.left {
  transform: translateX(-90px) rotateY(-90deg);
}
.top {
  transform: translateY(-90px) rotateX(90deg);
}
.bottom {
  transform: translateY(90px) rotateX(-90deg);
}

/* Unique gradient colors */
.front .sticker {
  background: linear-gradient(135deg, #ff1e56, #ff577f);
}
.back .sticker {
  background: linear-gradient(135deg, #ff9f1c, #ffbf69);
}
.right .sticker {
  background: linear-gradient(135deg, #2ecc71, #54e6a3);
}
.left .sticker {
  background: linear-gradient(135deg, #1e90ff, #4ab8ff);
}
.top .sticker {
  background: linear-gradient(135deg, #ecf0f1, #bdc3c7);
}
.bottom .sticker {
  background: linear-gradient(135deg, #f1c40f, #f4d03f);
}

/* Cube rotation with consistent square shape */
@keyframes rotate-cube {
  0% {
    transform: rotateX(35.264deg) rotateY(45deg) rotateZ(0deg);
  }
  25% {
    transform: rotateX(35.264deg) rotateY(135deg) rotateZ(0deg);
  }
  50% {
    transform: rotateX(35.264deg) rotateY(225deg) rotateZ(0deg);
  }
  75% {
    transform: rotateX(35.264deg) rotateY(315deg) rotateZ(0deg);
  }
  100% {
    transform: rotateX(35.264deg) rotateY(405deg) rotateZ(0deg);
  }
}

/* Sticker bounce animation */
@keyframes sticker-bounce {
  0% {
    transform: translateY(0) scale(1);
  }
  50% {
    transform: translateY(0) scale(1.04);
  }
  100% {
    transform: translateY(0) scale(1);
  }
}
</style>
`

  const loaders = [
    loader1, loader2, loader3, loader4, loader5, loader6,
    loader7, loader8, loader9, loader10, loader11, loader12
  ];
  const randomIndex = Math.floor(Math.random() * loaders.length);
  return loaders[randomIndex];
}
