<<<<<<< HEAD
/*
 * Programming Quiz: JuliaJames (4-1)
 * Loop through the numbers 1 to 20
 * If the number is divisible by 3, print "Julia"
 * If the number is divisible by 5, print "James"
 * If the number is divisible by 3 and 5, print "JuliaJames"
 * If the number is not divisible by 3 or 5, print the number
 */

for (x=1; x<21; x++) {
  if ((x%3==0) && (x%5==0)) {
    console.log("JuliaJames");
  } else if ((x%3==0) && !(x%5==0)) {
    console.log("Julia");
  } else if (!(x%3==0) && (x%5==0)) {
    console.log("James");
  } else {
    console.log(x);
  }
}
=======
/*
 * Programming Quiz: JuliaJames (4-1)
 * Loop through the numbers 1 to 20
 * If the number is divisible by 3, print "Julia"
 * If the number is divisible by 5, print "James"
 * If the number is divisible by 3 and 5, print "JuliaJames"
 * If the number is not divisible by 3 or 5, print the number
 */

for (x=1; x<21; x++) {
  if ((x%3==0) && (x%5==0)) {
    console.log("JuliaJames");
  } else if ((x%3==0) && !(x%5==0)) {
    console.log("Julia");
  } else if (!(x%3==0) && (x%5==0)) {
    console.log("James");
  } else {
    console.log(x);
  }
}
>>>>>>> 138fe44f08be31486d4b8f8bfa908688a9ee0163
