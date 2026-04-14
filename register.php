<?php 
include "db.php"; 

$message = "";
$messageType = ""; 
$shouldRedirect = false; 

if(isset($_POST['register'])){
    $user = $conn->real_escape_string($_POST['username']);
    $pass = password_hash($_POST['password'], PASSWORD_DEFAULT);

    // Check if username already exists
    $check = $conn->query("SELECT * FROM users WHERE username='$user'");

    if($check->num_rows > 0){
        $message = "Username already exists! Choose another one.";
        $messageType = "error";
    } else {
        // Insert new user
        if($conn->query("INSERT INTO users (username, password) VALUES ('$user','$pass')")){
            $message = "Registration successful! Redirecting to login...";
            $messageType = "success";
            $shouldRedirect = true; 
        } else {
            $message = "Error creating account. Please try again.";
            $messageType = "error";
        }
    }
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Register - 16th IMB Finance System</title>
    <style>
        :root {
            --primary-green: #1b5e20;
            --hover-green: #165030;
            --error-red: #d32f2f;
            --success-green: #2e7d32;
        }

        body { 
            margin: 0; 
            font-family: 'Segoe UI', Roboto, sans-serif; 
            background: #f0f2f5; 
            display: flex; 
            justify-content: center; 
            align-items: center; 
            height: 100vh; 
        }

        .register-card { 
            background: white; 
            padding: 40px; 
            border-radius: 12px; 
            box-shadow: 0 8px 24px rgba(0,0,0,0.15); 
            width: 100%;
            max-width: 380px; 
            text-align: center; 
            border-top: 6px solid var(--primary-green);
        }

        .register-card h2 { 
            margin: 10px 0 5px;
            color: var(--primary-green); 
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 1px;
        }

        .system-subtitle {
            font-size: 13px;
            color: #666;
            margin-bottom: 25px;
            display: block;
        }

        .input-group { 
            margin-bottom: 18px; 
            text-align: left; 
        }

        .input-group label { 
            display: block; 
            margin-bottom: 6px; 
            color: #444; 
            font-size: 14px; 
            font-weight: 600;
        }

        .input-group input { 
            width: 100%; 
            padding: 12px; 
            border: 1px solid #ccc; 
            border-radius: 6px; 
            box-sizing: border-box; 
            font-size: 16px;
            transition: 0.3s;
        }

        .input-group input:focus {
            border-color: var(--primary-green);
            outline: none;
            box-shadow: 0 0 0 3px rgba(27, 94, 32, 0.1);
        }

        .register-btn { 
            width: 100%; 
            padding: 13px; 
            background: var(--primary-green); 
            color: white; 
            border: none; 
            border-radius: 6px; 
            font-size: 16px;
            font-weight: bold; 
            cursor: pointer; 
            transition: 0.3s;
            margin-top: 10px;
        }

        .register-btn:hover { 
            background: var(--hover-green);
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        }

        .alert { 
            padding: 12px; 
            border-radius: 6px; 
            margin-bottom: 20px; 
            font-size: 14px; 
            text-align: left;
            border-left: 4px solid transparent;
        }

        .error { 
            background: #ffebee; 
            color: var(--error-red); 
            border-left-color: var(--error-red);
        }

        .success { 
            background: #e8f5e9; 
            color: var(--success-green); 
            border-left-color: var(--success-green);
        }

        .footer-link { 
            margin-top: 25px; 
            padding-top: 20px;
            border-top: 1px solid #eee;
            font-size: 14px; 
            color: #555; 
        }

        .footer-link a { 
            color: var(--primary-green); 
            text-decoration: none; 
            font-weight: 700; 
        }

        .footer-link a:hover { 
            text-decoration: underline; 
        }
    </style>

    <?php if($shouldRedirect): ?>
    <script>
        setTimeout(function(){
            window.location.href = "login.php";
        }, 2000);
    </script>
    <?php endif; ?>
</head>
<body>

<div class="register-card">
    <h2>System</h2>
    <span class="system-subtitle">16th IMB Procurement Management System Register</span>

    <?php if($message): ?>
        <div class="alert <?php echo $messageType; ?>">
            <?php echo $message; ?>
        </div>
    <?php endif; ?>

    <form method="POST">
        <div class="input-group">
            <label>Username</label>
            <input type="text" name="username" placeholder="Choose a username" required>
        </div>
        
        <div class="input-group">
            <label>Password</label>
            <input type="password" name="password" placeholder="Create a password" required>
        </div>

        <button type="submit" name="register" class="register-btn">CREATE ACCOUNT</button>
    </form>

    <div class="footer-link">
        Already have an account? <a href="login.php">Sign In Here</a>
    </div>
</div>

</body>
</html>