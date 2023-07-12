package com.kg;

import java.util.List;
import java.util.concurrent.*;
import java.util.logging.Logger;
import java.util.stream.Collectors;

import static java.util.stream.IntStream.range;

public class Rig {
    static Logger logger = Logger.getLogger(Rig.class.getName());
    public static void main(String[] args) {
        //Synchronous
       logger.info(XLSXToJson.asJsonString("testsheet.xlsx","Test Results"));

       //Parallel
        ExecutorService executorService = Executors.newFixedThreadPool(200, r -> {
            Thread t = Executors.defaultThreadFactory().newThread(r);
            t.setDaemon(true);
            return t;
        });

        List<Future<String>> futures = range(1,50).mapToObj(x -> {
            return executorService.submit(new Callable<String>() {
                @Override
                public String call() throws Exception {
                    //you are about to start <range> amount of io ops in your machine so know your machines limits
                    return XLSXToJson.asJsonString("testsheet.xlsx","Test Results");
                }
            });
        }).collect(Collectors.toList());

        //do other things here;

        futures.forEach(x -> {
            try {
                logger.info(x.get());
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            } catch (ExecutionException e) {
                throw new RuntimeException(e);
            }
        });

        executorService.shutdown();
    }
}