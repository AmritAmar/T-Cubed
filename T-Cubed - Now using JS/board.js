class Board {

  constructor(r, c) {
    this.row = r;
    this.col = c;
    this.active = false;
    this.moves = [
      [-1, -1, -1],
      [-1, -1, -1],
      [-1, -1, -1]
    ]
    this.winningSquares = [
      [-1, -1],
      [-1, -1],
      [-1, -1]
    ]
  }

  resetBoard() {
    for (let i = 0; i < 3; i++) {
      for (let j = 0; j < 3; j++) {
        this.moves[i][j] = -1;
      }
    }
  }

  play(r, c, val) {
    this.moves[r][c] = val;
  }

  isBoardSomeEmpty() {
    for (let i = 0; i < 3; i++) {
      for (let j = 0; j < 3; j++) {
        if (this.moves[i][j] == -1) return true;
      }
    }
    return false;
  }

  equals3(a, b, c) {
    return (a == b && b == c && a != -1);
  }

  isBoardConquered() {
    let winner = -1;

    // vertical
    for (let i = 0; i < 3; i++) {
      if (this.equals3(this.moves[i][0], this.moves[i][1], this.moves[i][2])) {
        winner = this.moves[i][0];
      }
    }

    // horizontal
    for (let i = 0; i < 3; i++) {
      if (this.equals3(this.moves[0][i], this.moves[1][i], this.moves[2][i])) {
        winner = this.moves[0][i];
      }
    }

    // Diagonal
    if (this.equals3(this.moves[0][0], this.moves[1][1], this.moves[2][2])) {
      winner = this.moves[0][0];
    }
    if (this.equals3(this.moves[2][0], this.moves[1][1], this.moves[0][2])) {
      winner = this.moves[2][0];
    }

    if (winner == -1 && this.isBoardSomeEmpty()) {
      return -1;
    } else if (winner == -1 && !this.isBoardSomeEmpty()) {
      return 0;
    }
    return winner;
  }

}