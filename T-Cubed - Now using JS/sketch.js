const P1 = 1;
const P2 = 2;

let GameBoard = [];
let CurrentPlayer = P1;
let special = false;

function setup() {
  createCanvas(600, 800);
  for (let i = 0; i < 3; i++) {
    let tempBoards = []
    for (let j = 0; j < 3; j++) {
      tempBoards.push(new Board(j + 1, i + 1));
    }
    GameBoard.push(tempBoards);
  }

  // set row 1, col 1 board as active
  GameBoard[1][1].active = true;
}

function getMouseOnCorrectBoard() {
  let row = floor(mouseY / 200);
  let col = floor(mouseX / 200);

  return GameBoard[row][col];
}

function getMouseOnSquare() {
  let sx = floor(mouseX / 200) * 200;
  let sy = floor(mouseY / 200) * 200;
  let nx = mouseX - sx;
  let ny = mouseY - sy;
  let row = floor(ny / (200 / 3));
  let col = floor(nx / (200 / 3));

  return [row, col];
}

function mousePressed() {
  if (mouseX > 600 || mouseY > 600) return;

  let correctBoard = getMouseOnCorrectBoard();
  if (!correctBoard.active) return;

  let correctSquare = getMouseOnSquare();
  if (correctBoard.moves[correctSquare[0]][correctSquare[1]] != -1) return;

  playMove(correctBoard, correctSquare);

}

function nextPlayer() {
  if (CurrentPlayer == P1) {
    CurrentPlayer = P2;
  } else {
    CurrentPlayer = P1;
  }
}

function playMove(boardToPlayOn, squareToPlayOn) {

  //Play the move
  boardToPlayOn.play(squareToPlayOn[0], squareToPlayOn[1], CurrentPlayer);

  //Switch Boards to the move the player moved to
  boardToPlayOn.active = false;

  if (GameBoard[squareToPlayOn[0]][squareToPlayOn[1]].isBoardSomeEmpty()) {

    for (var b1 of GameBoard) {
      for (var b2 of b1) {
        b2.active = false;
      }
    }
    special = false;

    GameBoard[squareToPlayOn[0]][squareToPlayOn[1]].active = true;

  } else {
    special = true;
    for (var b1 of GameBoard) {
      for (var b2 of b1) {
        b2.active = true;
      }
    }

  }

  nextPlayer();

}

function drawBoard(board) {

  //Get position of board based on row/col
  posX = board.row * 100 * 2 - 100;
  posY = board.col * 100 * 2 - 100;

  //Make appropriate adjustments
  translate(posX, posY);
  let newSize = (this.width / 3 - ((this.width / 3) / 100 * 2));
  textAlign(CENTER);
  textSize(newSize / 3);

  //Draw outer box
  if (board.active) {
    stroke(255, 0, 255);
  } else if (board.isBoardConquered() == 0) {
    stroke(255, 255, 255);
  } else if (board.isBoardConquered() == 1) {
    stroke(255, 0, 0);
  } else if (board.isBoardConquered() == 2) {
    stroke(0, 255, 255);
  } else {
    stroke(100);
  }
  noFill();
  square(-newSize / 2, -newSize / 2, newSize);

  //Draw the lines
  for (let i = 1; i <= 2; i++) {
    line(-newSize / 2,
      -newSize / 2 + i * newSize / 3,
      newSize / 2,
      -newSize / 2 + i * newSize / 3);
  }
  for (let i = 1; i <= 2; i++) {
    line(-newSize / 2 + i * newSize / 3,
      -newSize / 2,
      -newSize / 2 + i * newSize / 3,
      newSize / 2);
  }

  //Draw the moves within the board
  for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
      if (board.moves[i][j] == 1) {
        stroke(255, 0, 0);
        text('X',
          -newSize / 2 + (j * 2 + 1) * newSize / 6,
          -newSize / 2 + (i * 2 + 1) * newSize / 6 + newSize / 9);
      } else if (board.moves[i][j] == 2) {
        stroke(0, 255, 255);
        text('O',
          -newSize / 2 + (j * 2 + 1) * newSize / 6,
          -newSize / 2 + (i * 2 + 1) * newSize / 6 + newSize / 9);
      }
    }
  }

  translate(-posX, -posY);
}

function equals3(a, b, c) {
  return (a == b && b == c && a != -1);
}

function getGameBoardState() {
  let GameBoardState = [];
  for (let i = 0; i < 3; i++) {
    let tempBoards = []
    for (let j = 0; j < 3; j++) {
      tempBoards.push(GameBoard[i][j].isBoardConquered());
    }
    GameBoardState.push(tempBoards);
  }
  return GameBoardState;
}

function isGameOver() {
  let winner = -1;
  let GameBoardState = getGameBoardState();

  // vertical
  for (let i = 0; i < 3; i++) {
    if (equals3(GameBoardState[i][0], GameBoardState[i][1], GameBoardState[i][2])) {
      winner = GameBoardState[i][0];
    }
  }

  // horizontal
  for (let i = 0; i < 3; i++) {
    if (equals3(GameBoardState[0][i], GameBoardState[1][i], GameBoardState[2][i])) {
      winner = GameBoardState[0][i];
    }
  }

  // Diagonal
  if (equals3(GameBoardState[0][0], GameBoardState[1][1], GameBoardState[2][2])) {
    winner = GameBoardState[0][0];
  }
  if (equals3(GameBoardState[2][0], GameBoardState[1][1], GameBoardState[0][2])) {
    winner = GameBoardState[2][0];
  }

  return winner;
}

function draw() {
  background(30);

  // Instructions
  textSize(16);
  textAlign(CENTER);
  stroke(175);
  fill(255);
  let s = 'This is ULTIMATE Tic Tac Toe! \n Take turns playing on the highlighted board (purple). However, your move determines where your opponent will place their next move. \n The first player to conquer a board claims that board. However, that board can still be played in (even though no moves count, as it has already been won). The game ends when 3 boards are won in tic tac toe fashion. A move into a board that has no squares left to play in allows the player to play anywhere they want.';
  text(s, 10, 630, 570, 800)
  
  if (isGameOver() != -1) {
    //Game Won, set everything to false
    for (var b1 of GameBoard) {
      for (var b2 of b1) {
        b2.active = false;
      }
    }
    //Draw the board one final time
    for (var b1 of GameBoard) {
      for (var b2 of b1) {
        drawBoard(b2);
      }
    }
    noLoop();
  } else {
    for (var b1 of GameBoard) {
      for (var b2 of b1) {
        drawBoard(b2);
      }
    }
  }
}