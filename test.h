#ifndef TEST
#define TEST

#define MAX_PATH 260
int main(int argc, char **argv);
/*
* This is a multiline comment
* because, as you might know,
* C allows comments that can
* span many lines
* to exist in its code
*/

typedef struct _T {
  int   item1;
  int   item2;  //this is also a comment
  char *item3;
} SOMETHING;

typedef enum _H {
  hA,
  hB=3,
  hC=0x3ff,
  hD=0o167
} HELLO;

/***************************************
 The following is how you would find
 the definition of an API function in
 the Win32 Platform SDK Documentation
***************************************/
LRESULT SendMessage(
  HWND hWnd,      // handle to destination window
  UINT Msg,       // message
  WPARAM wParam,  // first message parameter
  LPARAM lParam   // second message parameter
);
#endif