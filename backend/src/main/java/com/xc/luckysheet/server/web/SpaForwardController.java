package com.xc.luckysheet.server.web;

import jakarta.servlet.http.HttpServletRequest;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.server.ResponseStatusException;

import static org.springframework.http.HttpStatus.NOT_FOUND;

/** Serves the browser application's index for extension-less deep links. */
@Controller
public class SpaForwardController {
    @RequestMapping({
            "/",
            "/{first:^(?!api$|ws$|assets$)[^\\.]+$}",
            "/{first:^(?!api$|ws$|assets$)[^\\.]+$}/{second:[^\\.]+}",
            "/{first:^(?!api$|ws$|assets$)[^\\.]+$}/{second:[^\\.]+}/{third:[^\\.]+}",
            "/{first:^(?!api$|ws$|assets$)[^\\.]+$}/{second:[^\\.]+}/{third:[^\\.]+}/{fourth:[^\\.]+}",
            "/{first:^(?!api$|ws$|assets$)[^\\.]+$}/{second:[^\\.]+}/{third:[^\\.]+}/{fourth:[^\\.]+}/{fifth:[^\\.]+}"
    })
    public String forward(HttpServletRequest request) {
        String path = request.getRequestURI();
        if (path.startsWith("/api/") || "/ws".equals(path) || path.startsWith("/assets/")) {
            throw new ResponseStatusException(NOT_FOUND);
        }
        return "forward:/index.html";
    }
}
